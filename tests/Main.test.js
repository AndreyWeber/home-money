const { createSandbox } = require('./setup/sandbox');

let ctx;
beforeEach(() => { ctx = createSandbox(); });

// ─────────────────────────────────────────────────────────────────────────────
describe('createCellNote', () => {
  const date = new Date(2024, 2, 8); // 3/8/2024

  it('returns existing note unchanged when noteToAdd is empty and isPlanned is false', () => {
    expect(ctx.createCellNote('existing note', date, 100, null, false)).toBe('existing note');
    expect(ctx.createCellNote('existing note', date, 100, '', false)).toBe('existing note');
    expect(ctx.createCellNote('existing note', date, 100, undefined, false)).toBe('existing note');
  });

  it('returns undefined note unchanged when both noteToAdd and isPlanned are falsy', () => {
    expect(ctx.createCellNote(undefined, date, 100, null, false)).toBe(undefined);
  });

  it('appends a regular note to empty existing note', () => {
    const result = ctx.createCellNote('', date, 150, 'Groceries', false);
    expect(result).toBe('3/8/2024: 150 - Groceries');
  });

  it('appends a regular note with line feed after existing note', () => {
    const result = ctx.createCellNote('previous line', date, 150, 'Groceries', false);
    expect(result).toBe('previous line\n3/8/2024: 150 - Groceries');
  });

  it('creates a planned note without a comment', () => {
    const result = ctx.createCellNote('', date, 200, null, true);
    expect(result).toBe('3/8/2024: 200 - потрачено');
  });

  it('creates a planned note with a comment', () => {
    const result = ctx.createCellNote('', date, 200, 'Rent', true);
    expect(result).toBe('3/8/2024: 200 - Rent - потрачено');
  });

  it('appends a planned note after existing note', () => {
    const result = ctx.createCellNote('line 1', date, 50, null, true);
    expect(result).toBe('line 1\n3/8/2024: 50 - потрачено');
  });

  it('appends a planned note with comment after existing note', () => {
    const result = ctx.createCellNote('line 1', date, 50, 'Transport', true);
    expect(result).toBe('line 1\n3/8/2024: 50 - Transport - потрачено');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('calculateCellNoteMoneySum', () => {
  it('returns 0 for null note', () => {
    expect(ctx.calculateCellNoteMoneySum(null)).toBe(0);
  });

  it('returns 0 for empty string note', () => {
    expect(ctx.calculateCellNoteMoneySum('')).toBe(0);
  });

  it('returns 0 for note with no numbers', () => {
    expect(ctx.calculateCellNoteMoneySum('no numbers here')).toBe(0);
  });

  it('sums a single numeric token', () => {
    expect(ctx.calculateCellNoteMoneySum('3/8/2024: 150 - Groceries')).toBe(150);
  });

  it('sums multiple numeric tokens across lines', () => {
    const note = '3/8/2024: 150 - Groceries\n3/9/2024: 75 - Transport';
    expect(ctx.calculateCellNoteMoneySum(note)).toBe(225);
  });

  it('handles decimal values', () => {
    const note = '3/8/2024: 99.99 - Item';
    expect(ctx.calculateCellNoteMoneySum(note)).toBeCloseTo(99.99);
  });

  it('ignores non-numeric tokens like dates (4/3/2024: is not fully numeric)', () => {
    // "4/3/2024:" – parseFloat gives 4, isFinite gives false for the full token
    // so the date token is skipped; only the standalone number counts
    const note = '4/3/2024: 50 - Note';
    expect(ctx.calculateCellNoteMoneySum(note)).toBe(50);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('calculateCellBackgroung', () => {
  it('returns noFill when cellNote is null', () => {
    expect(ctx.calculateCellBackgroung(null, 100)).toBe(ctx.CELL_BACKGROUND_COLOR.noFill);
  });

  it('returns noFill when cellNote is empty string', () => {
    expect(ctx.calculateCellBackgroung('', 100)).toBe(ctx.CELL_BACKGROUND_COLOR.noFill);
  });

  it('returns noFill when cellValue is 0 (falsy)', () => {
    expect(ctx.calculateCellBackgroung('3/8/2024: 100 - Groceries', 0))
      .toBe(ctx.CELL_BACKGROUND_COLOR.noFill);
  });

  it('returns noFill when cellValue is null', () => {
    expect(ctx.calculateCellBackgroung('3/8/2024: 100 - Groceries', null))
      .toBe(ctx.CELL_BACKGROUND_COLOR.noFill);
  });

  it('returns notSpent when note money sum is 0', () => {
    // Note has no numbers
    expect(ctx.calculateCellBackgroung('some note with no numbers', 100))
      .toBe(ctx.CELL_BACKGROUND_COLOR.notSpent);
  });

  it('returns almostSpent when note sum is less than cell value', () => {
    const note = '3/8/2024: 50 - Groceries';
    expect(ctx.calculateCellBackgroung(note, 100))
      .toBe(ctx.CELL_BACKGROUND_COLOR.almostSpent);
  });

  it('returns spent when note sum equals cell value', () => {
    const note = '3/8/2024: 100 - Groceries';
    expect(ctx.calculateCellBackgroung(note, 100))
      .toBe(ctx.CELL_BACKGROUND_COLOR.spent);
  });

  it('returns spent when note sum exceeds cell value', () => {
    const note = '3/8/2024: 150 - Groceries';
    expect(ctx.calculateCellBackgroung(note, 100))
      .toBe(ctx.CELL_BACKGROUND_COLOR.spent);
  });

  it('CELL_BACKGROUND_COLOR constants have expected hex values', () => {
    expect(ctx.CELL_BACKGROUND_COLOR.noFill).toBe('#ffffff');
    expect(ctx.CELL_BACKGROUND_COLOR.notSpent).toBe('#e6b8af');
    expect(ctx.CELL_BACKGROUND_COLOR.almostSpent).toBe('#d9ead3');
    expect(ctx.CELL_BACKGROUND_COLOR.spent).toBe('#c9daf8');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('processTransaction', () => {
  // Helper: builds a minimal raw data row
  function makeRow(overrides) {
    return Object.assign({
      dateOfTransaction: new Date(2020, 0, 1), // well in the past
      value: 100,
      symbol: 'Food',
      comment: 'Groceries',
      plannedPayment: false,
      timestamp: new Date(2020, 0, 1, 10, 0, 0),
    }, overrides);
  }

  // Helper: sets up summaryBalanceSheet mock for a symbol found at row 1
  function setupSheet(cellValue = 50, cellNote = '') {
    const { mockSheet, mockCell, mockRange } = ctx._mocks;

    // Row values for symbol lookup: first column
    mockSheet.getSheetValues.mockReturnValue([['Food'], ['Transport']]);
    mockSheet.getMaxRows.mockReturnValue(2);
    mockSheet.getMaxColumns.mockReturnValue(3);

    // Column lookup: getRange(rowNum, 1, 1, maxCols).getValues() — find first empty col
    const colRangeValues = [['Food', 50, '']]; // col 3 is empty → colNum = 2
    const colRange = {
      getValues: jest.fn().mockReturnValue(colRangeValues),
      getCell: jest.fn().mockReturnValue(mockCell),
    };
    mockSheet.getRange.mockImplementation((row, col, numRows, numCols) => {
      if (numRows === 1 && numCols !== undefined) return colRange;
      return { getCell: jest.fn().mockReturnValue(mockCell) };
    });

    mockCell.getValue.mockReturnValue(cellValue);
    mockCell.getNote.mockReturnValue(cellNote);

    return mockSheet;
  }

  it('returns false when transaction date is in the future', () => {
    const futureDate = new Date(Date.now() + 86400000 * 365);
    const row = makeRow({ dateOfTransaction: futureDate });
    expect(ctx.processTransaction(row, ctx._mocks.mockSheet)).toBe(false);
  });

  it('returns false when transaction value is 0', () => {
    const row = makeRow({ value: 0 });
    expect(ctx.processTransaction(row, ctx._mocks.mockSheet)).toBe(false);
  });

  it('returns false when transaction value is negative', () => {
    const row = makeRow({ value: -10 });
    expect(ctx.processTransaction(row, ctx._mocks.mockSheet)).toBe(false);
  });

  it('returns false when symbol is not found in the sheet', () => {
    ctx._mocks.mockSheet.getSheetValues.mockReturnValue([['Other']]);
    ctx._mocks.mockSheet.getMaxRows.mockReturnValue(1);
    const row = makeRow({ symbol: 'NonExistent' });
    expect(ctx.processTransaction(row, ctx._mocks.mockSheet)).toBe(false);
  });

  it('returns true and adds transaction value to cell for non-planned payment', () => {
    const sheet = setupSheet(50, '');
    const row = makeRow({ value: 100, comment: null, plannedPayment: false });

    const result = ctx.processTransaction(row, sheet);

    expect(result).toBe(true);
    expect(ctx._mocks.mockCell.setValue).toHaveBeenCalledWith(150); // 50 + 100
  });

  it('sets cell note for a non-planned transaction with comment', () => {
    const sheet = setupSheet(50, '');
    const row = makeRow({ value: 100, comment: 'Groceries', plannedPayment: false });

    ctx.processTransaction(row, sheet);

    expect(ctx._mocks.mockCell.setNote).toHaveBeenCalledWith(
      expect.stringContaining('Groceries')
    );
  });

  it('returns true and sets background for planned payment', () => {
    const sheet = setupSheet(200, '');
    const row = makeRow({ value: 150, comment: 'Rent', plannedPayment: true });

    const result = ctx.processTransaction(row, sheet);

    expect(result).toBe(true);
    expect(ctx._mocks.mockCell.setBackground).toHaveBeenCalled();
  });

  it('returns false and shows warning for planned payment when cell value is 0 and note sum > 0', () => {
    const existingNote = '1/1/2020: 100 - previous';
    const sheet = setupSheet(0, existingNote);
    const row = makeRow({ value: 50, comment: 'Rent', plannedPayment: true });

    const result = ctx.processTransaction(row, sheet);

    expect(result).toBe(false);
    // HtmlService should have been used to create the warning
    expect(ctx.HtmlService.createTemplateFromFile).toHaveBeenCalled();
  });
});
