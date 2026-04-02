const { createSandbox } = require('./setup/sandbox');

let ctx;
beforeEach(() => { ctx = createSandbox(); });

// Helper: configures the mockSheet so that getSheetValues() returns a
// column-vector matching the MetadataKeys ordering used in the real sheet.
function setupMetadataSheet(ctx) {
  const { mockSheet } = ctx._mocks;
  mockSheet.getSheetValues.mockReturnValue([
    ['LatestTransactionDate'],
    ['LatestTransactionName'],
    ['LatestTransactionSymbol'],
    ['LatestTransactionValue'],
  ]);
  mockSheet.getMaxRows.mockReturnValue(4);
  return mockSheet;
}

// Helper: configures the mockRange.getValues() to return a full metadata table
// (header row + data rows) that getRowsData / getMetadataObjects will consume.
function setupMetadataRangeData(ctx) {
  ctx._mocks.mockRange.getValues.mockReturnValue([
    ['key', 'value'],
    ['LatestTransactionDate',   '2024-03-08T00:00:00.000Z'],
    ['LatestTransactionName',   'Groceries'],
    ['LatestTransactionSymbol', 'Food'],
    ['LatestTransactionValue',  150],
  ]);
}

// ─────────────────────────────────────────────────────────────────────────────
describe('MetadataKeys', () => {
  it('exposes the expected constant keys', () => {
    expect(ctx.MetadataKeys.LATEST_TRANSACTION_DATE).toBe('LatestTransactionDate');
    expect(ctx.MetadataKeys.LATEST_TRANSACTION_NAME).toBe('LatestTransactionName');
    expect(ctx.MetadataKeys.LATEST_TRANSACTION_SYMBOL).toBe('LatestTransactionSymbol');
    expect(ctx.MetadataKeys.LATEST_TRANSACTION_VALUE).toBe('LatestTransactionValue');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('getMetadataSheet', () => {
  it('calls SpreadsheetApp to obtain the Metadata sheet', () => {
    ctx.getMetadataSheet();
    expect(ctx.SpreadsheetApp.getActiveSpreadsheet).toHaveBeenCalled();
    expect(ctx._mocks.mockSpreadsheet.getSheetByName).toHaveBeenCalledWith('Metadata');
  });

  it('returns the mocked sheet object', () => {
    expect(ctx.getMetadataSheet()).toBe(ctx._mocks.mockSheet);
  });

  it('caches the sheet on subsequent calls (SpreadsheetApp called only once)', () => {
    ctx.getMetadataSheet();
    ctx.getMetadataSheet();
    expect(ctx.SpreadsheetApp.getActiveSpreadsheet).toHaveBeenCalledTimes(1);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('getMetadataObjects', () => {
  it('returns an array of metadata objects from the sheet', () => {
    setupMetadataRangeData(ctx);

    const result = ctx.getMetadataObjects();

    expect(Array.isArray(result)).toBe(true);
    expect(result).toHaveLength(4);
    expect(result[0]).toEqual({ key: 'LatestTransactionDate', value: '2024-03-08T00:00:00.000Z' });
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('getMetadataObject', () => {
  it('throws when key is falsy', () => {
    expect(() => ctx.getMetadataObject(null)).toThrow();
    expect(() => ctx.getMetadataObject(undefined)).toThrow();
    expect(() => ctx.getMetadataObject('')).toThrow();
  });

  it('returns the metadata object matching the key', () => {
    setupMetadataRangeData(ctx);

    const result = ctx.getMetadataObject('LatestTransactionDate');

    expect(result).toEqual({
      key:   'LatestTransactionDate',
      value: '2024-03-08T00:00:00.000Z',
    });
  });

  it('returns undefined when key does not exist in metadata', () => {
    setupMetadataRangeData(ctx);

    expect(ctx.getMetadataObject('NonExistentKey')).toBeUndefined();
  });

  it('finds each known MetadataKey', () => {
    setupMetadataRangeData(ctx);

    expect(ctx.getMetadataObject(ctx.MetadataKeys.LATEST_TRANSACTION_DATE)).toBeDefined();
    expect(ctx.getMetadataObject(ctx.MetadataKeys.LATEST_TRANSACTION_NAME)).toBeDefined();
    expect(ctx.getMetadataObject(ctx.MetadataKeys.LATEST_TRANSACTION_SYMBOL)).toBeDefined();
    expect(ctx.getMetadataObject(ctx.MetadataKeys.LATEST_TRANSACTION_VALUE)).toBeDefined();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('getAllMetataObjectsJson', () => {
  it('returns a valid JSON string', () => {
    setupMetadataRangeData(ctx);

    const json = ctx.getAllMetataObjectsJson();

    expect(() => JSON.parse(json)).not.toThrow();
  });

  it('JSON contains all metadata entries', () => {
    setupMetadataRangeData(ctx);

    const parsed = JSON.parse(ctx.getAllMetataObjectsJson());

    expect(Array.isArray(parsed)).toBe(true);
    expect(parsed).toHaveLength(4);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('setMetadataValue', () => {
  it('calls getRange with the correct row and sets the value', () => {
    setupMetadataSheet(ctx);

    ctx.setMetadataValue('LatestTransactionDate', '2024-03-08T00:00:00.000Z');

    // LatestTransactionDate is at index 0 in the flat array → rowNum = 1
    expect(ctx._mocks.mockSheet.getRange).toHaveBeenCalledWith('B1');
    expect(ctx._mocks.mockRange.setValue).toHaveBeenCalledWith('2024-03-08T00:00:00.000Z');
  });

  it('writes to the correct row for the second key', () => {
    setupMetadataSheet(ctx);

    ctx.setMetadataValue('LatestTransactionName', 'Groceries');

    expect(ctx._mocks.mockSheet.getRange).toHaveBeenCalledWith('B2');
    expect(ctx._mocks.mockRange.setValue).toHaveBeenCalledWith('Groceries');
  });

  it('throws when the key does not exist in the sheet', () => {
    setupMetadataSheet(ctx);

    expect(() => ctx.setMetadataValue('NonExistentKey', 'val')).toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('setLatestTransactionDate', () => {
  it('throws when date is falsy', () => {
    expect(() => ctx.setLatestTransactionDate(null)).toThrow();
    expect(() => ctx.setLatestTransactionDate(undefined)).toThrow();
  });

  it('persists an ISO string to the metadata sheet', () => {
    setupMetadataSheet(ctx);
    const date = new Date(2024, 2, 8);

    ctx.setLatestTransactionDate(date);

    expect(ctx._mocks.mockRange.setValue).toHaveBeenCalledWith(date.toISOString());
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('setLatestTransactionName', () => {
  it('throws when name is not a string', () => {
    expect(() => ctx.setLatestTransactionName(42)).toThrow();
    expect(() => ctx.setLatestTransactionName(null)).toThrow();
    expect(() => ctx.setLatestTransactionName(undefined)).toThrow();
  });

  it('persists the name string to the metadata sheet', () => {
    setupMetadataSheet(ctx);

    ctx.setLatestTransactionName('Groceries');

    expect(ctx._mocks.mockRange.setValue).toHaveBeenCalledWith('Groceries');
  });

  it('accepts empty string', () => {
    setupMetadataSheet(ctx);

    expect(() => ctx.setLatestTransactionName('')).not.toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('setLatestTransactionSymbol', () => {
  it('throws when symbol is not a string', () => {
    expect(() => ctx.setLatestTransactionSymbol(42)).toThrow();
    expect(() => ctx.setLatestTransactionSymbol(null)).toThrow();
  });

  it('persists the symbol string to the metadata sheet', () => {
    setupMetadataSheet(ctx);

    ctx.setLatestTransactionSymbol('Food');

    expect(ctx._mocks.mockRange.setValue).toHaveBeenCalledWith('Food');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('setLatestTransactionValue', () => {
  it('throws when value is NaN', () => {
    expect(() => ctx.setLatestTransactionValue(NaN)).toThrow();
    expect(() => ctx.setLatestTransactionValue('not-a-number')).toThrow();
  });

  it('accepts numeric string (global isNaN coerces before checking)', () => {
    setupMetadataSheet(ctx);

    // global isNaN('150') is false, so this should NOT throw
    expect(() => ctx.setLatestTransactionValue('150')).not.toThrow();
  });

  it('persists a numeric value to the metadata sheet', () => {
    setupMetadataSheet(ctx);

    ctx.setLatestTransactionValue(150);

    expect(ctx._mocks.mockRange.setValue).toHaveBeenCalledWith(150);
  });

  it('accepts zero', () => {
    setupMetadataSheet(ctx);

    expect(() => ctx.setLatestTransactionValue(0)).not.toThrow();
  });

  it('accepts negative numbers', () => {
    setupMetadataSheet(ctx);

    expect(() => ctx.setLatestTransactionValue(-50)).not.toThrow();
  });
});
