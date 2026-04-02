# home-money
A Google App Scripts-powered application designed to enhance Google Sheets for efficient home finance management.

# Prerequisites
- `npm install -g @google/clasp` - read more about `clasp` [here](https://github.com/google/clasp)
- `npm install --save @types/google-apps-script` - for VS Code intellisense

# Unit Tests

## Setup

Install dependencies (only needed once):

```bash
npm install
```

## Running tests

Run all tests:

```bash
npm test
```

Watch mode (re-runs on file changes):

```bash
npx jest --watch
```

Run a single test file:

```bash
npx jest tests/HelperFunctions.test.js
```

## Test structure

| File | Covers | Tests |
|---|---|---|
| `tests/setup/sandbox.js` | Shared vm sandbox — loads all source files into a GAS-like context with mocked `SpreadsheetApp`, `Session`, `HtmlService` | — |
| `tests/HelperFunctions.test.js` | `normalizeHeader/s`, `isValidDate`, `stringIsEmpty`, `isString`, `isAlnum`, `isDigit`, `isNumber`, `isStrContainValidDate`, `flattenArray`, `findValueIndex`, `_throwErr`, `dateAsUtc`, `dateToFormattedString`, `dateTimeToFormattedString`, `toJsonString`, `getObjects`, `getRowsData` | 78 |
| `tests/Main.test.js` | `createCellNote`, `calculateCellNoteMoneySum`, `calculateCellBackgroung`, `processTransaction` | 47 |
| `tests/MetadataFunctions.test.js` | `MetadataKeys`, `getMetadataSheet`, `getMetadataObjects`, `getMetadataObject`, `getAllMetataObjectsJson`, `setMetadataValue`, `setLatestTransactionDate/Name/Symbol/Value` | 31 |
| `tests/RawDataFunctions.test.js` | `getRawDataSheet`, `getRawDataTransactionObjects`, `getRawDataTransactionObjectWithMaxDate` | 19 |
