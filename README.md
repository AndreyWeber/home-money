# home-money
A Google App Scripts-powered application designed to enhance Google Sheets for efficient home finance management.

# Prerequisites
- `npm install -g @google/clasp` - read more about `clasp` [here](https://github.com/google/clasp)
- `npm install --save @types/google-apps-script` - for VS Code intellisense

# Deployment

## How Clasp versioning works

Clasp separates two concepts:

- **Version** — an immutable, numbered snapshot of the script code (like a git tag). You create them manually; GAS never auto-increments them. Versions are the only thing a deployment can point to.
- **Deployment** — a published endpoint with a stable Deployment ID that points to a specific version. Updating a deployment to point to a newer version keeps the same ID, so callers (e.g. the Google Sheet's bound triggers) are unaffected.

The recommended workflow is:

```
Edit locally  →  push to GAS draft  →  freeze a version  →  deploy or update deployment
```

## Scripts

| Script | Command | What it does |
|---|---|---|
| `npm run push` | `clasp push` | Upload local files to the GAS working draft (HEAD). No version is created. |
| `npm run push:watch` | `clasp push --watch` | Same as above but re-pushes automatically on every file save. |
| `npm run version:create` | `clasp version` | Freeze the current HEAD into the next immutable version. Optionally pass a description: `npm run version:create -- "v1.2.0 - fix date parsing"`. |
| `npm run version:list` | `clasp versions` | List all versions with their numbers and descriptions. |
| `npm run deploy:list` | `clasp deployments` | List all deployments with their IDs and the version each points to. |
| `npm run deploy:new` | `clasp push && clasp version && clasp deploy` | **Full pipeline** — push, freeze a new version, then create a brand-new deployment pointing to it. Use this for the first release or when you want a second parallel deployment. |
| `npm run deploy:update` | `clasp push && clasp version && clasp deploy --deploymentId` | **Update existing deployment** — push, freeze a version, then move an existing deployment to it. Append the Deployment ID: `npm run deploy:update -- <deploymentId>`. The ID stays the same so no reconfiguration is needed in the Sheet. |

## Typical release flow

**First release** (creates a fresh deployment):
```bash
npm run deploy:new
```

**Subsequent releases** (updates the existing deployment — Deployment ID stays the same):
```bash
# Find your Deployment ID once
npm run deploy:list

# Then on every release
npm run deploy:update -- AKfycbxXXXXXXXXXXXXXXXX
```

**Push only** (for testing in the GAS editor without creating a version):
```bash
npm run push
```

> **Note:** `clasp` must be installed globally (`npm install -g @google/clasp`) and you must be logged in (`clasp login`) before running any of these scripts.

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
