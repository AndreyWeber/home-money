/**
 * @fileoverview Helper functions for reading spreadsheet data as arrays of
 * objects, and general-purpose utilities used across the project.
 */

/// <reference path="Types.js" />

// ── Spreadsheet data helpers ─────────────────────────────────────────────────

/**
 * Iterates row by row over a sheet's data range and returns an array of
 * objects, each keyed by the normalised column header name.
 *
 * Date cells are re-zoned to the script timezone so that the local date the
 * user typed into the sheet is preserved correctly, regardless of differences
 * between the sheet timezone and the script timezone.
 *
 * @param  {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet containing the data
 * @returns {Object[]} Array of row objects indexed by normalised header name
 */
function getRowsData(sheet) {
  if (!sheet) {
    _throwErr(`getRowsData failed with invalid args: ${JSON.stringify({ sheet })}`);
  }

  const data = sheet.getDataRange().getValues();
  const headers = normalizeHeaders(data[0]);
  const dataRows = data.slice(1);
  const scriptTimeZone = Session.getScriptTimeZone();

  return dataRows.map(row => {
    const object = {};
    row.forEach((val, colIdx) => {
      const key = headers[colIdx];
      if (val === undefined || stringIsEmpty(val)) {
        object[key] = null;
      } else if (val instanceof Date && isValidDate(val)) {
        // GAS reads sheet dates as UTC midnight, but the sheet timezone may be
        // ahead of the script timezone. Re-interpret the local wall-clock date
        // the user typed by keeping local time components in the script TZ.
        const dt = luxon.DateTime.fromJSDate(val).setZone(scriptTimeZone, { keepLocalTime: true });
        object[key] = dt.toJSDate();
      } else {
        object[key] = val;
      }
    });
    return object;
  });
}

/**
 * Maps each row in a 2-D array to an object whose field names come from keys.
 * Rows where every cell is an empty string are excluded from the result.
 *
 * @param  {Array<Array<*>>} data - 2-D array of cell values
 * @param  {string[]}        keys - Property names for the resulting objects
 * @returns {Object[]} Array of row objects
 */
function getObjects(data, keys) {
  return data.reduce((objects, row) => {
    const object = {};
    let hasData = false;
    row.forEach((cellData, j) => {
      if (stringIsEmpty(cellData)) return;
      object[keys[j]] = cellData;
      hasData = true;
    });
    if (hasData) objects.push(object);
    return objects;
  }, []);
}

/**
 * Normalises each string in an array of column headers.
 *
 * @param  {string[]} headers - Raw header strings to normalise
 * @returns {string[]} Array of normalised header strings
 */
function normalizeHeaders(headers) {
  return headers.map(normalizeHeader);
}

/**
 * Normalises a single header string to a camelCase JavaScript identifier.
 * Non-alphanumeric characters are stripped; leading digits are ignored so the
 * result always starts with a lowercase letter.
 *
 * @param  {string} header - Header string to normalise
 * @returns {string} camelCase identifier
 *
 * @example
 * normalizeHeader('First Name')            // → 'firstName'
 * normalizeHeader('Market Cap (millions)') // → 'marketCapMillions'
 * normalizeHeader('1 number at start')     // → 'numberAtStart'
 */
function normalizeHeader(header) {
  let key = '';
  let upperCase = false;
  for (const letter of header) {
    if (letter === ' ' && key.length > 0) { upperCase = true; continue; }
    if (!isAlnum(letter)) continue;
    if (key.length === 0 && isDigit(letter)) continue;
    key += upperCase ? letter.toUpperCase() : letter.toLowerCase();
    upperCase = false;
  }
  return key;
}

// ── Type-checking utilities ──────────────────────────────────────────────────

/**
 * Duck-type check for a valid {@link Date} instance.
 * Returns `true` only when val has a callable `getTime` that returns a finite number.
 *
 * @param  {*} val - Value to test
 * @returns {boolean} `true` if val is a valid Date
 */
const isValidDate = (val) =>
  Boolean(val) &&
  typeof val.getTime === 'function' &&
  !isNaN(val.getTime());

/**
 * Returns `true` if str is an empty string (`""`).
 *
 * @param  {*} str - Value to test
 * @returns {boolean}
 */
const stringIsEmpty = (str) => isString(str) && str === EMPTY_STRING;

/**
 * Returns `true` if str is of type `string`.
 *
 * @param  {*} str - Value to test
 * @returns {boolean}
 */
const isString = (str) => typeof str === 'string';

/**
 * Returns `true` if char is an ASCII alphanumeric character.
 *
 * @param  {string} char - Single character to test
 * @returns {boolean}
 */
const isAlnum = (char) =>
  (char >= 'A' && char <= 'Z') ||
  (char >= 'a' && char <= 'z') ||
  isDigit(char);

/**
 * Returns `true` if char is an ASCII decimal digit (`0`–`9`).
 *
 * @param  {string} char - Single character to test
 * @returns {boolean}
 */
const isDigit = (char) => char >= '0' && char <= '9';

/**
 * Returns `true` if n is a finite number or a string representation of one.
 *
 * @param  {*} n - Value to test
 * @returns {boolean}
 */
const isNumber = (n) => !isNaN(parseFloat(n)) && isFinite(n);

/**
 * Returns `true` if val is a non-numeric string that parses to a valid Date.
 *
 * @param  {*} val - Value to test
 * @returns {boolean}
 */
const isStrContainValidDate = (val) => !isNumber(val) && isValidDate(new Date(val));

// ── Array / search utilities ─────────────────────────────────────────────────

/**
 * Flattens a multi-dimensional array into a single-dimensional one.
 * Supports arbitrary nesting depth.
 *
 * @param  {Array} array - Array to flatten (may be nested)
 * @returns {Array} Flat array
 */
const flattenArray = (array) => array.flat(Infinity);

/**
 * Returns the 1-based index of the first element satisfying the predicate,
 * or `0` if no element matches.
 *
 * The 1-based convention mirrors the Google Sheets row/column numbering used
 * throughout the Spreadsheet API.
 *
 * @param  {Array}    dataArray      - Array to search
 * @param  {Function} isProperValue  - Predicate: `(value) => boolean`
 * @returns {number} 1-based index of the matching element, or `0` if not found
 */
function findValueIndex(dataArray, isProperValue) {
  const idx = dataArray.findIndex(isProperValue);
  return idx === -1 ? 0 : idx + 1;
}

// ── Error helper ─────────────────────────────────────────────────────────────

/**
 * Throws a new `Error` with the given message.
 * Centralising `throw` in an expression context lets it be used inside
 * ternary and short-circuit expressions.
 *
 * @param  {string} [msg] - Error message; defaults to a generic fallback
 * @throws {Error} Always throws
 */
const _throwErr = (msg) => { throw new Error(msg ?? 'Unexpected error occured'); };

// ── Date utilities ───────────────────────────────────────────────────────────

/**
 * Converts a Date to a UTC timestamp representing midnight of that calendar
 * day, enabling timezone-safe day-level comparisons.
 *
 * @param  {Date} dt - Date to convert
 * @returns {number} UTC timestamp (ms since epoch) for midnight of that day
 * @throws {Error} When dt is not a valid Date
 */
const dateAsUtc = (dt) => isValidDate(dt)
  ? Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate())
  : _throwErr("Can't convert date to UTC. Probably 'dt' argument is undefined");

/**
 * Formats a Date as `"M/D/YYYY"` (no zero-padding).
 *
 * @param  {Date} dt - Date to format
 * @returns {string} Formatted date string
 * @throws {Error} When dt is not a valid Date
 */
const dateToFormattedString = (dt) => isValidDate(dt)
  ? `${dt.getMonth() + 1}/${dt.getDate()}/${dt.getFullYear()}`
  : _throwErr("Can't format date as string. Probably 'dt' argument is undefined.");

/**
 * Formats a Date as `"M/D/YYYY H:M:S"` (no zero-padding).
 *
 * @param  {Date} dt - Date/time to format
 * @returns {string} Formatted date-time string
 * @throws {Error} When dt is not a valid Date
 */
const dateTimeToFormattedString = (dt) => isValidDate(dt)
  ? `${dt.getMonth() + 1}/${dt.getDate()}/${dt.getFullYear()} ${dt.getHours()}:${dt.getMinutes()}:${dt.getSeconds()}`
  : _throwErr("Can't format date/time as string. Probably 'dt' argument is undefined");

// ── Serialisation helpers ────────────────────────────────────────────────────

/**
 * Calls `callback(...callbackArgs)` and serialises the return value to a JSON
 * string.
 *
 * @param  {Function}  callback      - Function whose result will be serialised
 * @param  {Array|null} [callbackArgs=null] - Arguments to pass to callback
 * @returns {string} JSON string of the callback's return value
 * @throws {Error} When callback is falsy or callbackArgs is not an array
 */
function toJsonString(callback, callbackArgs = null) {
  if (!callback) {
    _throwErr("'callback' argument cannot be null or undefined");
  }
  const args = callbackArgs ?? [];
  if (!Array.isArray(args)) {
    _throwErr("'callbackArgs' argument must be array");
  }
  return JSON.stringify(callback(...args));
}
