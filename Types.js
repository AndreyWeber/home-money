/**
 * @fileoverview Shared JSDoc type definitions for the home-money Google Apps
 * Script project. This file contains no runtime code — only documentation
 * annotations used by IDEs and doc generators.
 */

/**
 * A single raw transaction record read from the 'Raw Data' sheet.
 *
 * @typedef  {Object}     RawTransaction
 * @property {Date|null}  dateOfTransaction - Date the transaction took place
 * @property {number}     value             - Transaction amount
 * @property {string}     symbol            - Budget category symbol
 * @property {string}     comment           - Transaction description
 * @property {boolean}    plannedPayment    - Whether this is a planned payment
 * @property {Date|null}  timestamp         - Row creation timestamp
 */

/**
 * A single key/value entry from the 'Metadata' sheet.
 *
 * @typedef  {Object}        MetadataEntry
 * @property {string}        key   - Metadata key name
 * @property {string|number} value - Metadata value
 */

/**
 * Result object returned by {@link getDailyExpensesData}.
 *
 * @typedef  {Object} DailyExpensesResult
 * @property {string} sumActual   - Actual expenses sum with currency suffix
 * @property {string} sumExpected - Expected expenses sum with currency suffix
 * @property {string} overrun     - Overrun amount, or "(нет)" if none
 */

/**
 * Result object returned by {@link getTransactionsHistoryData}.
 *
 * @typedef  {Object}      TransactionsHistoryResult
 * @property {boolean}     isSummaryBalanceSheet - Whether the active sheet is Summary Balance
 * @property {string|null} jsonCellA1Notation    - A1 notation of the active cell, or null
 * @property {string}      [json]                - Serialised transaction history JSON (present when isSummaryBalanceSheet is true)
 */
