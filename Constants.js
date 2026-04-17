/**
 * @fileoverview Global constants for the home-money Google Apps Script project.
 */

// ── Sheet name constants ─────────────────────────────────────────────────────

/** @enum {string} Names of all sheets used throughout the project */
const Sheets = {
  SUMMARY_BALANCE:      'Summary Balance',
  RAW_DATA:             'Raw Data',
  TRANSACTIONS_HISTORY: 'Transactions History',
  METADATA:             'Metadata',
};

// ── Cell background colour constants ─────────────────────────────────────────

/** @enum {string} Hex colour codes for planned-payment cell backgrounds */
const CELL_BACKGROUND_COLOR = {
  /** Cell budget fully or over spent */
  spent:       '#c9daf8',
  /** Cell budget partially spent */
  almostSpent: '#d9ead3',
  /** Cell budget not yet spent */
  notSpent:    '#e6b8af',
  /** No special fill */
  noFill:      '#ffffff',
};

// ── String constants ─────────────────────────────────────────────────────────

/** Column header value used to identify planned-payment rows */
const IS_TRANSACTION_PLANNED = 'Is planned';

/** Title of the planned-payment warning dialog */
const PLANNED_PAYMENT_WARNING_TITLE = 'Необходимо проверить список транзакций!';

/** Title of the daily expenses dialog */
const DAILY_EXPENSES_DATA_TITLE = 'Расходы за день';

/** Label text used to locate the free-daily-cash row on the Summary Balance sheet */
const FREE_DAILY_CASH_DATA_TITLE = 'Свободных наличных на день:';

/** Currency symbol appended to all monetary display strings */
const CURRENCY_SUFFIX = ' zł';

/** Canonical empty-string sentinel */
const EMPTY_STRING = '';
