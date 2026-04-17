/**
 * Creates a Node.js vm sandbox that mimics the Google Apps Script global
 * environment.  All source files are loaded into a single shared vm context
 * so that functions defined in one file can call functions defined in another,
 * exactly as they do inside GAS.
 *
 * Top-level `const` / `let` declarations are rewritten to `var` before
 * evaluation so that they become enumerable properties of the context object
 * (vm only hoists `var` and function declarations onto the context).
 */

const vm     = require('vm');
const fs     = require('fs');
const path   = require('path');
const luxon  = require('luxon');

const ROOT = path.resolve(__dirname, '..', '..');

function loadFile(context, filename) {
  const raw  = fs.readFileSync(path.join(ROOT, filename), 'utf8');
  // Replace only top-level (column-0) const/let with var so declarations
  // become context properties while indented block-scoped ones are left alone.
  const code = raw.replace(/^(const|let)\s/gm, 'var ');
  vm.runInContext(code, context);
}

/**
 * Factory – call once per test (in beforeEach) to get a clean context with
 * all source files loaded and all GAS globals mocked.
 *
 * Exposed mock objects live at `ctx._mocks` so individual tests can configure
 * return values and inspect calls.
 */
function createSandbox() {
  // ── Granular mock hierarchy ──────────────────────────────────────────────

  const mockCell = {
    getValue:      jest.fn().mockReturnValue(0),
    setValue:      jest.fn(),
    getNote:       jest.fn().mockReturnValue(''),
    setNote:       jest.fn(),
    setBackground: jest.fn(),
    getA1Notation: jest.fn().mockReturnValue('A1'),
  };

  const mockRange = {
    getValues: jest.fn().mockReturnValue([]),
    getValue:  jest.fn().mockReturnValue(''),
    setValue:  jest.fn(),
    getCell:   jest.fn().mockReturnValue(mockCell),
    getA1Notation: jest.fn().mockReturnValue('A1'),
  };

  const mockSheet = {
    getDataRange:   jest.fn().mockReturnValue(mockRange),
    getRange:       jest.fn().mockReturnValue(mockRange),
    getSheetValues: jest.fn().mockReturnValue([[]]),
    getMaxRows:     jest.fn().mockReturnValue(10),
    getMaxColumns:  jest.fn().mockReturnValue(5),
    deleteRow:      jest.fn(),
    getName:        jest.fn().mockReturnValue('Sheet1'),
    getActiveCell:  jest.fn().mockReturnValue(mockCell),
  };

  const mockSpreadsheet = {
    getSheetByName: jest.fn().mockReturnValue(mockSheet),
    getActiveSheet: jest.fn().mockReturnValue(mockSheet),
    show:           jest.fn(),
  };

  const mockUi = {
    createMenu:      jest.fn().mockReturnThis(),
    addItem:         jest.fn().mockReturnThis(),
    addToUi:         jest.fn(),
    showModalDialog: jest.fn(),
    showSidebar:     jest.fn(),
  };

  const mockHtmlOutput = {
    setSandboxMode: jest.fn().mockReturnThis(),
    setTitle:       jest.fn().mockReturnThis(),
    setHeight:      jest.fn().mockReturnThis(),
    setWidth:       jest.fn().mockReturnThis(),
  };

  const mockTemplate = {
    evaluate: jest.fn().mockReturnValue(mockHtmlOutput),
    data: null,
  };

  // ── vm context (= the GAS global scope) ─────────────────────────────────

  const context = {
    // JS built-ins required by source code
    console,
    JSON,
    Date,
    Array,
    Object,
    Math,
    Number,
    String,
    Boolean,
    RegExp,
    Error,
    TypeError,
    isNaN,
    isFinite,
    parseInt,
    parseFloat,
    undefined,
    NaN,
    Infinity,

    // GAS globals
    SpreadsheetApp: {
      getActiveSpreadsheet: jest.fn().mockReturnValue(mockSpreadsheet),
      getUi:                jest.fn().mockReturnValue(mockUi),
      SandboxMode:          { IFRAME: 'IFRAME' },
    },

    Session: {
      getScriptTimeZone:   jest.fn().mockReturnValue('Europe/London'),
      getActiveUserLocale: jest.fn().mockReturnValue('en-GB'),
    },

    HtmlService: {
      createHtmlOutputFromFile: jest.fn().mockReturnValue(mockHtmlOutput),
      createTemplateFromFile:   jest.fn().mockReturnValue(mockTemplate),
      SandboxMode:              { IFRAME: 'IFRAME' },
    },

    // Third-party libraries available as GAS globals
    luxon,

    // Expose mock handles so tests can configure return values / assert calls
    _mocks: {
      mockCell,
      mockRange,
      mockSheet,
      mockSpreadsheet,
      mockUi,
      mockHtmlOutput,
      mockTemplate,
    },
  };

  vm.createContext(context);

  // Load source files in dependency order
  loadFile(context, 'Types.js');
  loadFile(context, 'Constants.js');
  loadFile(context, 'HelperFunctions.js');
  loadFile(context, 'RawDataFunctions.js');
  loadFile(context, 'MetadataFunctions.js');
  loadFile(context, 'TransactionsHistoryFunctions.js');
  loadFile(context, 'Main.js');

  return context;
}

module.exports = { createSandbox };
