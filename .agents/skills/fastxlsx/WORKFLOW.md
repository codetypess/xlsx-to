# XLSX Workflow

Use this document as this skill's workflow for `.xlsx` tasks.

## CLI Entry Selection

Use the first available CLI entry:

```bash
fastxlsx <subcommand> ...
```

If `fastxlsx` is not on `PATH` but the package is available:

```bash
npx fastxlsx <subcommand> ...
```

Only when working inside the `fastxlsx` repository root:

```bash
npm run cli -- <subcommand> ...
```

This order keeps the workflow portable:

- `fastxlsx` covers globally installed or published-package usage.
- `npx fastxlsx` covers project-local package usage without requiring a global install.
- `npm run cli --` is last because it only works inside this repository root and depends on repo-local development tooling.

The examples below use `fastxlsx` as shorthand. Replace it with the available entry above.

## Default Workflow

1. Inspect before writing so sheet names, header rows, and workbook structure are confirmed first.
2. Read the exact sheet, cell, or table row that will change so structured rows and offsets are not guessed.
3. Apply the smallest fitting command so the change stays reviewable and roundtrip-safe.
4. Validate after writing so packaging or serialization regressions are caught immediately.
5. Re-read exact results when output correctness matters so the final workbook, not the intended command, is being verified.

Start with:

```bash
fastxlsx inspect path/to/file.xlsx
fastxlsx get path/to/file.xlsx --sheet Sheet1 --cell B2
fastxlsx validate path/to/file.xlsx
```

`inspect` now returns a `recommendedRead` object per sheet. Follow that routing hint before choosing between `sheet`, `config-table`, and `table`, and prefer reusing `recommendedRead.commands.*` directly instead of rebuilding the command by hand.

Use `--in-place` only when the user clearly wants to overwrite the source workbook.

## Hard Routing Rules

Keep these rules ahead of guesswork:

- Do not abandon this skill after the first failed, empty, or obviously misaligned read. Re-route within the CLI first.
- After `inspect`, prefer the target sheet's `recommendedRead.commands.*` output over hand-built read commands. Only override it when the user explicitly provides different row boundaries or a different sheet target.
- Do not start with `table` for an arbitrary sheet unless a profile already exists or header and data row boundaries are known.
- Do not treat `config-table` as the generic reader for every config-like sheet. Use it when one header row directly names the fields being edited.
- If `inspect` shows marker headers such as `@config` or `@define;...`, row 1 is not the business header row. Re-run discovery with another header row or generate a table profile.
- If a read returns structure rows such as `int`, `string`, `auto`, `>>`, `!!!`, `###`, or `-`, you are not looking at final business rows yet.

## Read Recovery

When the first read does not return the rows the user asked for, follow this order:

1. Run `fastxlsx inspect path/to/file.xlsx` and confirm the sheet name, previewed headers, and `recommendedRead` route.
2. If `recommendedRead.commands.list` or `recommendedRead.commands.inspect` matches the user's target, run that suggested command first instead of constructing a new one.
3. If row 1 looks like a marker row rather than real field names, re-run discovery with another preview row:

```bash
fastxlsx inspect path/to/file.xlsx --header-row 2
```

4. If `sheet records list` or `sheet export` returns type rows or structure markers instead of business rows, switch to `table` with explicit boundaries or a generated profile:

```bash
fastxlsx table generate-profiles path/to/file.xlsx
fastxlsx table inspect path/to/file.xlsx --sheet Sheet1 --header-row 1 --data-start-row 6
fastxlsx table list path/to/file.xlsx --profile 'file#Sheet1'
```

5. If `table ... --profile` fails because the profile does not exist yet, generate profiles and retry. If profile inference still cannot find the sheet, fall back to explicit `--header-row` and `--data-start-row`.
6. Re-read the exact row or cell after switching commands so the final answer is based on workbook output, not on the intended command.

Treat "wrong rows returned" as a routing problem, not as proof that the CLI cannot answer the question.

## Command Choice

Use `inspect` and `get` for read-only discovery:

```bash
fastxlsx inspect path/to/file.xlsx
fastxlsx get path/to/file.xlsx --sheet Sheet1 --cell B2
```

Use plain sheet read commands for ordinary `.xlsx` sheet content. For non-profile workbooks, this is the default way to read a sheet unless the user has identified a structured table layout:

```bash
fastxlsx sheet records list path/to/file.xlsx --sheet Data --header-row 1
fastxlsx sheet export path/to/file.xlsx --sheet Data --format json
fastxlsx sheet export path/to/file.xlsx --sheet Data --format csv --output rows.csv
```

Do not use `table inspect`, `table list`, or `table get` merely to read an arbitrary worksheet. `table` is for structured sheets where a profile already exists or the header row, data start row, and key fields are explicitly known.

Plain `sheet` reads fit best when row 1 is the real header row and business data starts immediately underneath it. If the first returned records look like schema rows, comments, validators, or sentinels, switch to `table`.

Use `set` for single-cell edits:

```bash
fastxlsx set path/to/file.xlsx --sheet Sheet1 --cell B2 --text "hello" --output out.xlsx
fastxlsx set path/to/file.xlsx --sheet Sheet1 --cell C2 --formula "B2*0.9" --cached-value 110.7 --output out.xlsx
fastxlsx set path/to/file.xlsx --sheet Sheet1 --cell D2 --clear --output out.xlsx
```

Use direct workbook and style commands for targeted changes:

```bash
fastxlsx add-sheet path/to/file.xlsx --sheet Summary --output out.xlsx
fastxlsx rename-sheet path/to/file.xlsx --from Sheet1 --to Config --output out.xlsx
fastxlsx delete-sheet path/to/file.xlsx --sheet Scratch --output out.xlsx
fastxlsx copy-style path/to/file.xlsx --sheet Config --from B2 --to C2 --output out.xlsx
fastxlsx set-number-format path/to/file.xlsx --sheet Config --cell B2 --format '0.00%' --output out.xlsx
fastxlsx set-background-color path/to/file.xlsx --sheet Config --cell B2 --color FFFF0000 --output out.xlsx
```

Use `apply` for deterministic multi-step edits:

```bash
fastxlsx apply path/to/file.xlsx --ops /tmp/fastxlsx-ops.json --output out.xlsx
```

Use `apply` only when the change genuinely spans multiple actions. Single-cell or single-command edits are easier to audit when kept as direct CLI commands.

For `apply --ops`, read [OPS-SCHEMA.md](OPS-SCHEMA.md).

Use `sheet import` and `sheet records` for plain header-mapped sheet writes:

```bash
fastxlsx sheet import path/to/file.xlsx --sheet Data --format json --from rows.json --mode update --key-field id --output out.xlsx
fastxlsx sheet records update path/to/file.xlsx --sheet Data --key-field id --value 1001 --record '{"desc":"patched"}' --output out.xlsx
fastxlsx sheet records upsert path/to/file.xlsx --sheet Data --key-field id --record '{"id":1001,"desc":"complete row"}' --output out.xlsx
```

Use `update` for partial matched-row edits. It only writes fields present in the input record and preserves omitted fields.

Use `upsert` only when the payload contains the full row you want after the command. On matched rows, `upsert` replaces the row and clears omitted fields. On missing rows, `upsert` inserts a new row.

Use `replace` only when the whole record set or table body should be rewritten.

Use `config-table` for simple header-based config sheets where one header row directly names the fields and the data rows begin immediately below it:

```bash
fastxlsx config-table init path/to/file.xlsx --sheet Config --headers '["Key","Value"]' --output out.xlsx
fastxlsx config-table list path/to/file.xlsx --sheet Config
fastxlsx config-table get path/to/file.xlsx --sheet Config --field Key --text timeout
fastxlsx config-table update path/to/file.xlsx --sheet Config --field Key --text timeout --record '{"Value":"30"}' --output out.xlsx
fastxlsx config-table upsert path/to/file.xlsx --sheet Config --field Key --record '{"Key":"timeout","Value":"30"}' --in-place
fastxlsx config-table delete path/to/file.xlsx --sheet Config --field Key --text timeout --output out.xlsx
fastxlsx config-table replace path/to/file.xlsx --sheet Config --records '[{"Key":"timeout","Value":"30"}]' --output out.xlsx
fastxlsx config-table sync path/to/file.xlsx --sheet Config --from-json config.json --mode update --output out.xlsx
fastxlsx config-table sync path/to/file.xlsx --sheet Config --from-json config.json --mode upsert --output out.xlsx
```

If `inspect` shows marker headers such as `@config`, or the actual config fields only appear on a later row, do not keep forcing `config-table` with the default header row. Re-run discovery with another `--header-row`, or generate a profile and use `table --profile` instead.

Use `table` for structured sheets with explicit header and data row boundaries, not as the generic sheet reader:

```bash
fastxlsx table inspect path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6
fastxlsx table list path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6
fastxlsx table get path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key 1001 --key-field id
fastxlsx table update path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key-field id --key 1001 --record '{"desc":"patched"}' --output out.xlsx
fastxlsx table upsert path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key-field id --record '{"id":1001,"desc":"..."}' --in-place
fastxlsx table delete path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key 1001 --key-field id --output out.xlsx
fastxlsx table sync path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key-field id --from-json rows.json --mode update --output out.xlsx
fastxlsx table sync path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key-field id --from-json rows.json --mode replace --output out.xlsx
```

Treat rows such as `auto`, `>>`, `!!!`, `###`, and `-` as structure to preserve, not built-in business semantics.

These marker rows are a strong signal that `table` is the right reader, usually with a `dataStartRow` after the marker block.

## Profiles

If `table-profiles.json` already exists, prefer `--profile`. The default profiles file is `table-profiles.json`; override it with `--profiles-file` when needed.

Profiles matter because they freeze the agreed sheet name, header row, data start row, and key fields in one place instead of repeating them in every command.

```bash
fastxlsx table list res/task.xlsx --profile 'task#main'
fastxlsx table get res/task.xlsx --profile 'task#conf' --key '"GATE_SIEGE_TIME"'
fastxlsx table get res/task.xlsx --profile 'task#define' --key '{"key1":"TASK_TYPE","key2":"MAIN"}'
fastxlsx table update res/task.xlsx --profile 'task#main' --key 1001 --record '{"desc":"patched"}' --in-place
fastxlsx table upsert res/task.xlsx --profile 'task#main' --record '{"id":1001,"desc":"..."}' --in-place
fastxlsx table sync res/task.xlsx --profile 'task#conf' --from-json conf.json --mode update --output out.xlsx
fastxlsx table sync res/task.xlsx --profile 'task#conf' --from-json conf.json --mode upsert --output out.xlsx
```

When a workbook uses metadata rows, sentinel rows, composite keys, or non-row-1 headers, prefer generating or reusing profiles early. This is often the fastest way to get clean reads.

If profiles do not exist yet, generate them first:

```bash
fastxlsx table generate-profiles res/task.xlsx
fastxlsx table generate-profiles res/task.xlsx res/monster.xlsx --sheet-filter '^(main|conf)$' --output table-profiles.json
```

Workbook open failures, sheets whose table profile cannot be inferred, and duplicate generated profile names are skipped. The JSON output includes `skipped` entries with `file`, `reason`, and optional `sheet` or `profileName`.

When `--output` is used, stdout only prints `Generated profile file: <path>`; read the output file for the full generated `profiles` object and any `skipped` records.

For large workbook sets, avoid passing every path as a command argument. Write the paths to a newline-delimited file and use `--files-from` so shell argument length limits are not hit:

```bash
find res -name '*.xlsx' > /tmp/fastxlsx-files.txt
fastxlsx table generate-profiles --files-from /tmp/fastxlsx-files.txt --output table-profiles.json
```

You can also scan a directory recursively and ignore specific workbooks:

```bash
fastxlsx table generate-profiles --from-dir res --ignore res/archive/old.xlsx --output table-profiles.json
```

Generated names use `文件名#表名`, for example `task#main`.

## Limits

Prefer the CLI over ad hoc scripts or direct workbook XML edits.

That constraint is deliberate: the CLI is the reviewed surface that preserves roundtrip behavior. Direct XML edits or throwaway scripts bypass the guardrails this skill is trying to enforce.

If the current CLI cannot express the requested change:

1. Confirm that existing commands are insufficient.
2. Extend the CLI in this repository.
3. Re-run the workbook change through the CLI.
