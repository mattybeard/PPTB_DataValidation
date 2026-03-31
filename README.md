# DataValidation

DataValidation is a Power Platform ToolBox tool for running simple, repeatable data quality checks against Dataverse tables.

The tool lets you select a table, choose columns, configure tests per column, and run those tests across all records to produce a headline pass rate and per-test results.

## Current Capabilities

### Table and Column Setup

- Browse Dataverse tables (display name + logical name).
- Load table columns and filter out unsupported/internal types.
- Sort columns by logical name, display name, type, or configured test count.
- Configure tests per column or apply tests in bulk to selected columns.

### Implemented Tests

- `Contains Data`
  - Passes when the value is not null, undefined, or empty/whitespace.

- `Matches Regex`
  - Uses a configured regular expression against the string form of the value.

- `Matches Metadata` (currently implemented for selected attribute types)
  - `StringType`: value length must be `<= MaxLength`.
  - `IntegerType`: value must be an integer within `[MinValue, MaxValue]` (where bounds exist).
  - `DecimalType` and `MoneyType`: numeric value must be within `[MinValue, MaxValue]` (where bounds exist).
  - `PicklistType`, `StateType`, `StatusType`: value must exist in metadata options.

## How Test Execution Works

When you click `Run Selected Tests`, the tool:

1. Builds a FetchXML query including only columns with configured tests.
2. Retrieves all records for the selected table (paged).
3. Executes configured tests against each record value.
4. Displays:
   - A headline pass rate across all executed checks.
   - A per-column, per-test pass/fail breakdown.

## Local Development

### Prerequisites

- Node.js 18+

### Install

```bash
npm install
```

or

```bash
pnpm install
```

### Run Dev Mode

```bash
npm run dev
```

### Build

```bash
npm run build
```

### Validate Package

```bash
npm run validate
```

## Usage in Power Platform ToolBox

1. Build the tool (`npm run build`).
2. Package/install in ToolBox using your normal ToolBox workflow.
3. Connect to a Dataverse environment.
4. Select a table.
5. Configure tests on one or more columns.
6. Run tests and review the summary/results grid.

## Notes and Limitations

- `Matches Metadata` is intentionally incremental and currently covers the types listed above.
- Empty values are treated as metadata-valid so that completeness checks remain the job of `Contains Data`.
- Very large tables may take longer due to full-table paging and in-memory evaluation.

## Tech Stack

- React 18 + TypeScript
- MobX
- Fluent UI React v8
- Vite
- `window.dataverseAPI` and `window.toolboxAPI` from Power Platform ToolBox

## License

MIT
