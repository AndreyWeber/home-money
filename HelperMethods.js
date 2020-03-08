// ===================================================================
// Helper methods required to read spreadsheet data as arrays of
// objects.
// ===================================================================

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first row
//       or all the cells below columnHeadersRowIndex (if defined).
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  var headersIndex = columnHeadersRowIndex || range ? range.getRowIndex() - 1 : 1;
  var dataRange = range ||
    sheet.getRange(headersIndex + 1, 1, sheet.getMaxRows() - headersIndex, sheet.getMaxColumns());
  var numColumns = dataRange.getEndColumn() - dataRange.getColumn() + 1;
  var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(dataRange.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Empty Strings are returned for all Strings that could not be successfully normalized.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    keys.push(normalizeHeader(headers[i]));
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

// ===================================================================
// Helper methods for transactions processing.
// ===================================================================

// Converts multi-dimensional array to a single-dimensional one.
// Arguments:
//  - array: array to flatten
// Returns array.
function flattenArray(array) {
  var result = [];
  function traverse(arr) {
    for (var i = 0; i < arr.length; i++) {
      if (arr[i] instanceof Array) {
        traverse(arr[i]);
      } else {
        result.push(arr[i]);
      }
    }
  }
  traverse(array);
  
  return result;
}

// Searches and returns index of value which fits to isProperValue predicate in data array
// Arguments:
//  - dataArray: array of spreadsheet data to search for
//  - isProperValue: predicate to fit data array value for
// Returns: integer
function findValueIndex(dataArray, isProperValue) {
  function iter(idx) {
    if (dataArray.length == 0) {
      return 0;
    }
    
    return isProperValue(dataArray.shift()) ? idx : iter(idx + 1);
  }
  
  return iter(1);
}

// isNumber checks if n is number or string representation of number
// Arguments:
//   - n: value to test
// Returns boolean.
function isNumber(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

// Convert date to UTC format.
// Argumetns:
//   - dt: date to convert
// Returns date in UTC format.
function dateAsUtc(dt) {
  if (!dt) {
    throw new Error("Parameter dt is undefined.");
  }
  
  return Date.UTC(dt.getFullYear(), dt.getMonth() + 1, dt.getDate());
}

// Convert date to formatted string by next pattern "mm/dd/YYYY"
// Arguments:
//   - dt: date to convert
// Returns string.
function dateToFormattedString(dt) {
  if (!dt) {
    throw new Error("Parameter dt is undefined.");
  }
  
  return dt.getMonth() + 1 + "/" + dt.getDate() + "/" + dt.getFullYear();
}

function dateTimeToFormattedString(dt) {
  if (!dt) {
    throw new Error("Parameter dt is undefined.");
  }
  
  return dt.getMonth() + 1 + "/" + dt.getDate() + "/" + dt.getFullYear() + " " +
    dt.getHours() + ":" + dt.getMinutes() + ":" + dt.getSeconds();
}

// stringIsNotEmpty checks if input string is not empty
// Arguments:
//   - v: input string to check
// Returns string
function stringIsNotEmpty(v) {
  if (typeof(v) != "string") {
    throw new TypeError();
  }
  
  return v === "" ? false : true;
  
}

// sprintf prints formatted string where all entries of %s placeholder will be substituted with
// one of the provided parameters.
// Arguments:
//   - format: string to format
//   - etc: parameters which will substitute all %s placeholders in format string
// Returns string.
function sprintf(format, etc) {
    var arg = arguments;
    var i = 1;
    return format.replace(/%((%)|s)/g, function (m) { return m[2] || arg[i++] })
}