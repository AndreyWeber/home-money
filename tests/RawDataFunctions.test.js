const { createSandbox } = require('./setup/sandbox');

let ctx;
beforeEach(() => { ctx = createSandbox(); });

// ─────────────────────────────────────────────────────────────────────────────
describe('getRawDataSheet', () => {
  it('calls SpreadsheetApp to obtain the Raw Data sheet', () => {
    ctx.getRawDataSheet();
    expect(ctx.SpreadsheetApp.getActiveSpreadsheet).toHaveBeenCalled();
    expect(ctx._mocks.mockSpreadsheet.getSheetByName).toHaveBeenCalledWith('Raw Data');
  });

  it('returns the mocked sheet object', () => {
    expect(ctx.getRawDataSheet()).toBe(ctx._mocks.mockSheet);
  });

  it('caches the sheet (SpreadsheetApp called only once on repeated calls)', () => {
    ctx.getRawDataSheet();
    ctx.getRawDataSheet();
    expect(ctx.SpreadsheetApp.getActiveSpreadsheet).toHaveBeenCalledTimes(1);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('getRawDataTransactionObjects', () => {
  it('returns an array of transaction objects read from the raw data sheet', () => {
    ctx._mocks.mockRange.getValues.mockReturnValue([
      ['dateOfTransaction', 'value', 'symbol', 'comment', 'plannedPayment', 'timestamp'],
      [new Date(2024, 2, 8), 100, 'Food', 'Groceries', false, new Date(2024, 2, 8, 10, 0, 0)],
      [new Date(2024, 2, 9), 200, 'Transport', 'Bus', false, new Date(2024, 2, 9, 8, 0, 0)],
    ]);

    const result = ctx.getRawDataTransactionObjects();

    expect(Array.isArray(result)).toBe(true);
    expect(result).toHaveLength(2);
    expect(result[0]).toMatchObject({ value: 100, symbol: 'Food', comment: 'Groceries' });
    expect(result[1]).toMatchObject({ value: 200, symbol: 'Transport' });
  });

  it('returns empty array when sheet has only a header row', () => {
    ctx._mocks.mockRange.getValues.mockReturnValue([
      ['dateOfTransaction', 'value', 'symbol', 'comment', 'plannedPayment', 'timestamp'],
    ]);

    expect(ctx.getRawDataTransactionObjects()).toEqual([]);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('getRawDataTransactionObjectWithMaxDate', () => {
  it('throws when argument is falsy', () => {
    expect(() => ctx.getRawDataTransactionObjectWithMaxDate(null)).toThrow();
    expect(() => ctx.getRawDataTransactionObjectWithMaxDate(undefined)).toThrow();
  });

  it('returns the initial-value object for an empty array', () => {
    const result = ctx.getRawDataTransactionObjectWithMaxDate([]);
    // reduce with no valid items returns the initial value
    expect(result).toMatchObject({ comment: '', symbol: '' });
    expect(ctx.isValidDate(result.dateOfTransaction)).toBe(false); // NaN date
  });

  it('returns the only element for a single-element array', () => {
    const date = new Date(2024, 2, 8);
    const row = { dateOfTransaction: date, comment: 'Groceries', symbol: 'Food', value: 100 };

    const result = ctx.getRawDataTransactionObjectWithMaxDate([row]);

    expect(result).toBe(row);
  });

  it('returns the element with the latest transaction date', () => {
    const earlier = { dateOfTransaction: new Date(2024, 0, 1), comment: 'A', symbol: 'X', value: 10 };
    const later   = { dateOfTransaction: new Date(2024, 5, 15), comment: 'B', symbol: 'Y', value: 20 };
    const latest  = { dateOfTransaction: new Date(2024, 11, 31), comment: 'C', symbol: 'Z', value: 30 };

    const result = ctx.getRawDataTransactionObjectWithMaxDate([earlier, latest, later]);

    expect(result).toBe(latest);
  });

  it('ignores entries with missing dateOfTransaction', () => {
    const noDate  = { dateOfTransaction: null, comment: 'X', symbol: 'X', value: 10 };
    const withDate = { dateOfTransaction: new Date(2024, 2, 8), comment: 'Y', symbol: 'Y', value: 20 };

    const result = ctx.getRawDataTransactionObjectWithMaxDate([noDate, withDate]);

    expect(result).toBe(withDate);
  });

  it('ignores entries with invalid Date objects', () => {
    const invalid  = { dateOfTransaction: new Date('bad'), comment: 'X', symbol: 'X', value: 10 };
    const valid    = { dateOfTransaction: new Date(2024, 2, 8), comment: 'Y', symbol: 'Y', value: 20 };

    const result = ctx.getRawDataTransactionObjectWithMaxDate([invalid, valid]);

    expect(result).toBe(valid);
  });

  it('returns initial-value object when all entries have invalid dates', () => {
    const rows = [
      { dateOfTransaction: null,             comment: 'A', symbol: 'X', value: 1 },
      { dateOfTransaction: new Date('bad'),  comment: 'B', symbol: 'Y', value: 2 },
    ];

    const result = ctx.getRawDataTransactionObjectWithMaxDate(rows);

    expect(result).toMatchObject({ comment: '', symbol: '' });
  });

  it('handles an array with two items having the same date by returning one of them', () => {
    const d = new Date(2024, 2, 8);
    const row1 = { dateOfTransaction: new Date(d), comment: 'A', symbol: 'X', value: 1 };
    const row2 = { dateOfTransaction: new Date(d), comment: 'B', symbol: 'Y', value: 2 };

    const result = ctx.getRawDataTransactionObjectWithMaxDate([row1, row2]);

    // Either is acceptable; just verify it's one of them
    expect([row1, row2]).toContain(result);
  });
});
