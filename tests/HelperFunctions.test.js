const { createSandbox } = require('./setup/sandbox');

let ctx;
beforeEach(() => { ctx = createSandbox(); });

// ─────────────────────────────────────────────────────────────────────────────
describe('normalizeHeader', () => {
  it('lowercases a single word', () => {
    expect(ctx.normalizeHeader('Hello')).toBe('hello');
  });

  it('converts two-word header to camelCase', () => {
    expect(ctx.normalizeHeader('First Name')).toBe('firstName');
  });

  it('handles multi-word header', () => {
    expect(ctx.normalizeHeader('Date Of Transaction')).toBe('dateOfTransaction');
  });

  it('strips non-alphanumeric characters', () => {
    expect(ctx.normalizeHeader('Market Cap (millions)')).toBe('marketCapMillions');
  });

  it('ignores leading digit so output starts with a letter', () => {
    expect(ctx.normalizeHeader('1 number at the beginning is ignored'))
      .toBe('numberAtTheBeginningIsIgnored');
  });

  it('returns empty string for empty input', () => {
    expect(ctx.normalizeHeader('')).toBe('');
  });

  it('returns empty string for digits-only input', () => {
    expect(ctx.normalizeHeader('123')).toBe('');
  });

  it('strips special characters entirely', () => {
    expect(ctx.normalizeHeader('---')).toBe('');
  });

  it('handles header with colons and slashes', () => {
    expect(ctx.normalizeHeader('Date: mm/dd')).toBe('dateMmdd');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('normalizeHeaders', () => {
  it('returns empty array for empty input', () => {
    expect(ctx.normalizeHeaders([])).toEqual([]);
  });

  it('normalizes each header in the array', () => {
    expect(ctx.normalizeHeaders(['First Name', 'Date Of Transaction', 'Value']))
      .toEqual(['firstName', 'dateOfTransaction', 'value']);
  });

  it('produces empty strings for un-normalizable headers', () => {
    expect(ctx.normalizeHeaders(['123', '---'])).toEqual(['', '']);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('isValidDate', () => {
  it('returns true for a valid Date object', () => {
    expect(ctx.isValidDate(new Date(2024, 2, 8))).toBe(true);
  });

  it('returns true for Date constructed from ISO string', () => {
    expect(ctx.isValidDate(new Date('2024-03-08'))).toBe(true);
  });

  it('returns false for Invalid Date', () => {
    expect(ctx.isValidDate(new Date('not-a-date'))).toBe(false);
  });

  it('returns false for NaN Date', () => {
    expect(ctx.isValidDate(new Date(NaN))).toBe(false);
  });

  it('returns falsy for null', () => {
    expect(ctx.isValidDate(null)).toBeFalsy();
  });

  it('returns falsy for undefined', () => {
    expect(ctx.isValidDate(undefined)).toBeFalsy();
  });

  it('returns falsy for a string (no getTime method)', () => {
    expect(ctx.isValidDate('2024-03-08')).toBeFalsy();
  });

  it('returns falsy for a number (no getTime method)', () => {
    expect(ctx.isValidDate(42)).toBeFalsy();
  });

  it('returns falsy for a plain object (no getTime method)', () => {
    expect(ctx.isValidDate({})).toBeFalsy();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('stringIsEmpty', () => {
  it('returns true for empty string', () => {
    expect(ctx.stringIsEmpty('')).toBe(true);
  });

  it('returns false for non-empty string', () => {
    expect(ctx.stringIsEmpty('hello')).toBe(false);
  });

  it('returns false for null', () => {
    expect(ctx.stringIsEmpty(null)).toBe(false);
  });

  it('returns false for undefined', () => {
    expect(ctx.stringIsEmpty(undefined)).toBe(false);
  });

  it('returns false for number 0', () => {
    expect(ctx.stringIsEmpty(0)).toBe(false);
  });

  it('returns false for whitespace string', () => {
    expect(ctx.stringIsEmpty(' ')).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('isString', () => {
  it('returns true for string literal', () => {
    expect(ctx.isString('hello')).toBe(true);
  });

  it('returns true for empty string', () => {
    expect(ctx.isString('')).toBe(true);
  });

  it('returns false for number', () => {
    expect(ctx.isString(42)).toBe(false);
  });

  it('returns false for null', () => {
    expect(ctx.isString(null)).toBe(false);
  });

  it('returns false for undefined', () => {
    expect(ctx.isString(undefined)).toBe(false);
  });

  it('returns false for boolean', () => {
    expect(ctx.isString(true)).toBe(false);
  });

  it('returns false for object', () => {
    expect(ctx.isString({})).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('isAlnum', () => {
  it('returns true for lowercase letter', () => {
    expect(ctx.isAlnum('a')).toBe(true);
  });

  it('returns true for uppercase letter', () => {
    expect(ctx.isAlnum('Z')).toBe(true);
  });

  it('returns true for digit', () => {
    expect(ctx.isAlnum('5')).toBe(true);
  });

  it('returns false for space', () => {
    expect(ctx.isAlnum(' ')).toBe(false);
  });

  it('returns false for punctuation', () => {
    expect(ctx.isAlnum('!')).toBe(false);
  });

  it('returns false for underscore', () => {
    expect(ctx.isAlnum('_')).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('isDigit', () => {
  it('returns true for "0"', () => {
    expect(ctx.isDigit('0')).toBe(true);
  });

  it('returns true for "9"', () => {
    expect(ctx.isDigit('9')).toBe(true);
  });

  it('returns false for letter', () => {
    expect(ctx.isDigit('a')).toBe(false);
  });

  it('returns false for space', () => {
    expect(ctx.isDigit(' ')).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('isNumber', () => {
  it('returns true for integer', () => {
    expect(ctx.isNumber(42)).toBe(true);
  });

  it('returns true for float', () => {
    expect(ctx.isNumber(3.14)).toBe(true);
  });

  it('returns true for negative number', () => {
    expect(ctx.isNumber(-7)).toBe(true);
  });

  it('returns true for numeric string', () => {
    expect(ctx.isNumber('42')).toBe(true);
  });

  it('returns true for float string', () => {
    expect(ctx.isNumber('3.14')).toBe(true);
  });

  it('returns false for empty string', () => {
    expect(ctx.isNumber('')).toBe(false);
  });

  it('returns false for non-numeric string', () => {
    expect(ctx.isNumber('hello')).toBe(false);
  });

  it('returns false for null', () => {
    expect(ctx.isNumber(null)).toBe(false);
  });

  it('returns false for NaN', () => {
    expect(ctx.isNumber(NaN)).toBe(false);
  });

  it('returns false for Infinity', () => {
    expect(ctx.isNumber(Infinity)).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('isStrContainValidDate', () => {
  it('returns true for ISO date string', () => {
    expect(ctx.isStrContainValidDate('2024-03-08')).toBe(true);
  });

  it('returns false for numeric string (isNumber is true)', () => {
    expect(ctx.isStrContainValidDate('42')).toBe(false);
  });

  it('returns false for non-date string', () => {
    expect(ctx.isStrContainValidDate('hello')).toBe(false);
  });

  it('returns false for number', () => {
    expect(ctx.isStrContainValidDate(42)).toBe(false);
  });

  it('returns false for empty string', () => {
    expect(ctx.isStrContainValidDate('')).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('flattenArray', () => {
  it('returns empty array unchanged', () => {
    expect(ctx.flattenArray([])).toEqual([]);
  });

  it('returns flat array unchanged', () => {
    expect(ctx.flattenArray([1, 2, 3])).toEqual([1, 2, 3]);
  });

  it('flattens one level of nesting', () => {
    expect(ctx.flattenArray([[1, 2], [3, 4]])).toEqual([1, 2, 3, 4]);
  });

  it('flattens deeply nested arrays', () => {
    expect(ctx.flattenArray([[1, [2, 3]], [4]])).toEqual([1, 2, 3, 4]);
  });

  it('flattens a column-vector (2D array with single-element rows)', () => {
    expect(ctx.flattenArray([['a'], ['b'], ['c']])).toEqual(['a', 'b', 'c']);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('findValueIndex', () => {
  it('returns 0 for empty array', () => {
    expect(ctx.findValueIndex([], () => true)).toBe(0);
  });

  it('returns 1 when first element matches', () => {
    expect(ctx.findValueIndex(['a', 'b', 'c'], v => v === 'a')).toBe(1);
  });

  it('returns 2 when second element matches', () => {
    expect(ctx.findValueIndex(['a', 'b', 'c'], v => v === 'b')).toBe(2);
  });

  it('returns 3 when last element matches', () => {
    expect(ctx.findValueIndex(['a', 'b', 'c'], v => v === 'c')).toBe(3);
  });

  it('returns 0 when no element matches', () => {
    expect(ctx.findValueIndex(['a', 'b', 'c'], v => v === 'z')).toBe(0);
  });

  it('works with numeric predicate', () => {
    expect(ctx.findValueIndex([10, 20, 30], v => v === 20)).toBe(2);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('_throwErr', () => {
  it('throws an Error with the provided message', () => {
    expect(() => ctx._throwErr('boom')).toThrow('boom');
  });

  it('throws an Error with default message when none provided', () => {
    expect(() => ctx._throwErr()).toThrow('Unexpected error occured');
  });

  it('throws an instance of Error', () => {
    expect(() => ctx._throwErr('x')).toThrow(Error);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('dateAsUtc', () => {
  it('returns a number for a valid date', () => {
    const result = ctx.dateAsUtc(new Date(2024, 2, 8));
    expect(typeof result).toBe('number');
  });

  it('returns the same value for two equal dates', () => {
    expect(ctx.dateAsUtc(new Date(2024, 2, 8)))
      .toBe(ctx.dateAsUtc(new Date(2024, 2, 8)));
  });

  it('returns a larger value for a later date', () => {
    expect(ctx.dateAsUtc(new Date(2024, 2, 9)))
      .toBeGreaterThan(ctx.dateAsUtc(new Date(2024, 2, 8)));
  });

  it('throws for an invalid Date', () => {
    expect(() => ctx.dateAsUtc(new Date('invalid'))).toThrow();
  });

  it('throws for null', () => {
    expect(() => ctx.dateAsUtc(null)).toThrow();
  });

  it('throws for undefined', () => {
    expect(() => ctx.dateAsUtc(undefined)).toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('dateToFormattedString', () => {
  it('formats date as m/d/YYYY', () => {
    expect(ctx.dateToFormattedString(new Date(2024, 2, 8))).toBe('3/8/2024');
  });

  it('includes full month and day without zero-padding', () => {
    expect(ctx.dateToFormattedString(new Date(2024, 11, 31))).toBe('12/31/2024');
  });

  it('formats January correctly (month index 0 → month number 1)', () => {
    expect(ctx.dateToFormattedString(new Date(2024, 0, 1))).toBe('1/1/2024');
  });

  it('throws for invalid Date', () => {
    expect(() => ctx.dateToFormattedString(new Date('bad'))).toThrow();
  });

  it('throws for null', () => {
    expect(() => ctx.dateToFormattedString(null)).toThrow();
  });

  it('throws for undefined', () => {
    expect(() => ctx.dateToFormattedString(undefined)).toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('dateTimeToFormattedString', () => {
  it('formats date-time as m/d/YYYY H:M:S', () => {
    expect(ctx.dateTimeToFormattedString(new Date(2024, 2, 8, 14, 30, 45)))
      .toBe('3/8/2024 14:30:45');
  });

  it('formats midnight correctly', () => {
    expect(ctx.dateTimeToFormattedString(new Date(2024, 0, 1, 0, 0, 0)))
      .toBe('1/1/2024 0:0:0');
  });

  it('throws for invalid Date', () => {
    expect(() => ctx.dateTimeToFormattedString(new Date('bad'))).toThrow();
  });

  it('throws for null', () => {
    expect(() => ctx.dateTimeToFormattedString(null)).toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('toJsonString', () => {
  it('throws when callback is null', () => {
    expect(() => ctx.toJsonString(null)).toThrow();
  });

  it('throws when callback is undefined', () => {
    expect(() => ctx.toJsonString(undefined)).toThrow();
  });

  it('serialises the return value of the callback', () => {
    const result = ctx.toJsonString(() => ({ key: 'value' }));
    expect(result).toBe('{"key":"value"}');
  });

  it('passes callbackArgs to the callback', () => {
    const fn = (a, b) => a + b;
    expect(ctx.toJsonString(fn, [2, 3])).toBe('5');
  });

  it('throws when callbackArgs is not an array', () => {
    expect(() => ctx.toJsonString(() => {}, 'not-array')).toThrow();
  });

  it('accepts null callbackArgs (treated as empty array)', () => {
    expect(() => ctx.toJsonString(() => 'ok', null)).not.toThrow();
  });

  it('serialises array return values', () => {
    expect(ctx.toJsonString(() => [1, 2, 3])).toBe('[1,2,3]');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('getObjects', () => {
  it('returns empty array when data is empty', () => {
    expect(ctx.getObjects([], ['key', 'value'])).toEqual([]);
  });

  it('maps row arrays to objects using provided keys', () => {
    const result = ctx.getObjects(
      [['Alice', 30], ['Bob', 25]],
      ['name', 'age']
    );
    expect(result).toEqual([{ name: 'Alice', age: 30 }, { name: 'Bob', age: 25 }]);
  });

  it('skips rows where all cells are empty strings', () => {
    const result = ctx.getObjects([['', ''], ['Alice', 30]], ['name', 'age']);
    expect(result).toEqual([{ name: 'Alice', age: 30 }]);
  });

  it('omits empty-string cell values from the resulting object', () => {
    const result = ctx.getObjects([['Alice', '']], ['name', 'note']);
    expect(result).toEqual([{ name: 'Alice' }]);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
describe('getRowsData', () => {
  it('throws when sheet is falsy', () => {
    expect(() => ctx.getRowsData(null)).toThrow();
    expect(() => ctx.getRowsData(undefined)).toThrow();
  });

  it('returns an array of objects indexed by normalised column names', () => {
    ctx._mocks.mockRange.getValues.mockReturnValue([
      ['Date Of Transaction', 'Value', 'Comment'],
      [new Date(2024, 2, 8), 100, 'Groceries'],
      [new Date(2024, 2, 9), 200, 'Transport'],
    ]);

    const result = ctx.getRowsData(ctx._mocks.mockSheet);

    expect(result).toHaveLength(2);
    expect(result[0]).toMatchObject({ value: 100, comment: 'Groceries' });
    expect(result[1]).toMatchObject({ value: 200, comment: 'Transport' });
  });

  it('converts empty-string cell values to null', () => {
    ctx._mocks.mockRange.getValues.mockReturnValue([
      ['Symbol', 'Comment'],
      ['Food', ''],
    ]);

    const result = ctx.getRowsData(ctx._mocks.mockSheet);
    expect(result[0]).toEqual({ symbol: 'Food', comment: null });
  });

  it('returns empty array when there are no data rows (only headers)', () => {
    ctx._mocks.mockRange.getValues.mockReturnValue([
      ['Symbol', 'Value'],
    ]);

    expect(ctx.getRowsData(ctx._mocks.mockSheet)).toEqual([]);
  });
});
