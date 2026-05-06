import * as xlsx from "fastxlsx";
import { basename } from "path";
import type { CheckerType } from "./contracts";
import { BuiltinChecker } from "./contracts";
import { convertValue } from "./conversion";
import { assert, error, trace } from "./errors";
import { DEFAULT_TAG, DEFAULT_WRITER, processors, writers } from "./registry";
import { type Field, type Sheet, type TCell, type TRow, Type } from "./schema";
import { checkType, ignoreField, toString } from "./value";
import { getContext } from "./workbook";

const MAX_HEADERS = 6;

const toLocation = (col: number, row: number) => {
    const COLUMN = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    let ret = "";
    col = col - 1;
    while (true) {
        const c = col % 26;
        ret = COLUMN[c] + ret;
        col = (col - c) / 26 - 1;
        if (col < 0) {
            break;
        }
    }
    return `${ret}${row}`;
};

const parseProcessor = (str: string) => {
    return str
        .split(/[;\n\r]+/)
        .map((s) => s.trim())
        .filter((s) => s)
        .map((s) => {
            const match = s.match(/^@(\w+)(?:\((.*?)\))?$/);
            const [, name = "", args = ""] = match ?? [];
            if (!name) {
                error(`Parse processor error: '${s}'`);
            } else if (!processors[name]) {
                error(`Processor not found: '${s}'`);
            }
            return {
                name,
                args: args
                    ? Array.from(args.matchAll(/{[^{}]+}|\[[^[\]]+\]|[^,]+/g)).map((a) =>
                          a[0].trim()
                      )
                    : [],
            };
        })
        .filter((p) => p.name);
};

const makeFilePath = (path: string) => (path.endsWith(".xlsx") ? path : path + ".xlsx");

export const parseChecker = (
    rowFile: string,
    rowSheet: string,
    location: string,
    index: number,
    str: string
): CheckerType[] => {
    if (str === "x" || (index === 0 && str.startsWith("!!"))) {
        return [];
    }
    if (str.trim() === "") {
        error(`No checker defined at ${location}`);
    }
    return str
        .split(/[;\n\r]+/)
        .map((s) => s.trim())
        .filter((s) => s)
        .map((s) => {
            const force = s.startsWith("!");
            if (force) {
                s = s.slice(1);
            }
            using _ = trace(`Parsing checker at ${location}: '${s}'`);
            let checker: CheckerType | undefined;
            if (s.startsWith("@")) {
                const [, name = "", arg = ""] = s.match(/^@(\w+)(?:\((.*?)\))?$/) ?? [];
                const args = arg.split(",").map((v) => v.trim());
                const oneof: CheckerType[] =
                    name === "oneof"
                        ? args.map((child) => {
                              if (!child) {
                                  error(`Invalid oneof checker at ${location}: '${s}'`);
                              }
                              const parsed: CheckerType[] = parseChecker(
                                  rowFile,
                                  rowSheet,
                                  location,
                                  index,
                                  child
                              );
                              if (parsed.length !== 1) {
                                  error(
                                      `Oneof branch must contain exactly one checker: '${child}'`
                                  );
                              }
                              return parsed[0]!;
                          })
                        : [];
                checker = {
                    name: name === "oneof" ? BuiltinChecker.OneOf : name,
                    force,
                    source: s,
                    location,
                    args,
                    oneof,
                    refers: {},
                    exec: null!,
                };
            } else if (s.startsWith("[") && s.endsWith("]")) {
                checker = {
                    name: BuiltinChecker.Range,
                    force,
                    source: s,
                    location,
                    args: [s],
                    oneof: [],
                    refers: {},
                    exec: null!,
                };
            } else if (s.endsWith("#")) {
                const [, rowKey = "", rowFilter = "", colFile = ""] =
                    s.match(/^(?:\$([^&]*)?(?:&(.+))?==)?([^#]*)#$/) ?? [];
                checker = {
                    name: BuiltinChecker.Sheet,
                    force,
                    source: s,
                    location,
                    args: [rowFile, rowSheet, rowKey, rowFilter, makeFilePath(colFile || rowFile)],
                    oneof: [],
                    refers: {},
                    exec: null!,
                };
            } else if (s.includes("#")) {
                const [
                    ,
                    rowKey = "",
                    rowFilter = "",
                    colFile = "",
                    colSheet = "",
                    colKey = "",
                    colFilter = "",
                ] = s.match(/^(?:\$([^&]*)?(?:&(.+))?==)?([^#=]*)#([^.]+)\.(\w+)(?:&(.+))?$/) ?? [];
                if (!colSheet || !colKey) {
                    error(`Invalid index checker at ${location}: '${s}'`);
                }
                checker = {
                    name: BuiltinChecker.Index,
                    force,
                    source: s,
                    location,
                    args: [
                        rowFile,
                        rowSheet,
                        rowKey,
                        rowFilter,
                        makeFilePath(colFile || rowFile),
                        colSheet,
                        colKey,
                        colFilter,
                    ],
                    oneof: [],
                    refers: {},
                    exec: null!,
                };
            } else if (s !== "x") {
                checker = {
                    name: BuiltinChecker.Expr,
                    force,
                    source: s,
                    location,
                    args: [s],
                    oneof: [],
                    refers: {},
                    exec: null!,
                };
            }
            return checker;
        })
        .filter((v): v is CheckerType => !!v);
};

const readCell = (sheet: xlsx.Sheet, r: number, c: number) => {
    const value = sheet.getCell(r, c);
    const v = typeof value === "string" ? value.trim() : (value ?? "");
    const cell: TCell = {
        v: v,
        r: toLocation(c, r),
        s: v.toString(),
    };
    cell["!type"] = Type.Cell;
    return cell;
};

export const loadHeader = (path: string, data: xlsx.Workbook) => {
    const ctx = getContext(DEFAULT_WRITER, DEFAULT_TAG)!;
    const requiredProcessors = Object.values(processors)
        .filter((p) => p.option.required)
        .reduce(
            (acc, p) => {
                acc[p.name] = 0;
                return acc;
            },
            {} as Record<string, number>
        );

    const workbook = ctx.get(path);
    const writerKeys = Object.keys(writers);

    let firstSheet: Sheet | null = null;

    for (const rawSheet of data.getSheets()) {
        using _ = trace(`Reading sheet '${rawSheet.name}' in '${path}'`);
        const firstCell = rawSheet.getCell(1, 1);
        if (rawSheet.name.startsWith("#") || !firstCell) {
            continue;
        }

        if (!rawSheet.name.match(/^[\w_]+$/)) {
            error(`Invalid sheet name: '${rawSheet}', only 'A-Za-z0-9_' are allowed`);
        }

        const sheet: Sheet = {
            name: rawSheet.name,
            ignore: false,
            processors: [],
            fields: [],
            data: {},
        };

        sheet.data["!type"] = Type.Sheet;
        sheet.data["!name"] = rawSheet.name;

        const str = firstCell.toString().trim();
        const colCount = rawSheet.columnCount;
        let r = 1;
        if (str.startsWith("@")) {
            sheet.processors.push(...parseProcessor(str));
            r = 2;
            for (const p of sheet.processors) {
                if (requiredProcessors[p.name] !== undefined) {
                    requiredProcessors[p.name]++;
                }
            }
        }

        if (!rawSheet.getCell(r, 1)) {
            continue;
        }

        const parsed: Record<string, boolean> = {};
        for (let c = 1; c <= colCount; c++) {
            const name = toString(readCell(rawSheet, r + 0, c));
            const typename = toString(readCell(rawSheet, r + 1, c));
            const writer = toString(readCell(rawSheet, r + 2, c));
            const checker = toString(readCell(rawSheet, r + 3, c));
            const comment = toString(readCell(rawSheet, r + 4, c));

            if (name && typename && writer !== "x") {
                const arr = writer
                    .split("|")
                    .map((w) => w.trim())
                    .filter((w) => c > 1 || !w.startsWith(">>"))
                    .filter((w) => w)
                    .map((w) => {
                        if (!writerKeys.includes(w)) {
                            error(`Writer not found: '${w}' at ${toLocation(c, r + 2)}`);
                        }
                        return w;
                    });
                if (parsed[name]) {
                    error(`Duplicate field name: '${name}' at ${toLocation(c, r)}`);
                }
                if (name.match(/[^A-Za-z0-9_$-]/)) {
                    error(
                        `Invalid field name: '${name}' at ${toLocation(c, r)}, only '$A-Za-z0-9_' are allowed`
                    );
                }
                parsed[name] = true;
                sheet.fields.push({
                    index: c - 1,
                    name,
                    typename,
                    writers: arr.length ? arr : writerKeys.slice(),
                    checkers: parseChecker(
                        basename(path),
                        rawSheet.name,
                        toLocation(c, r + 3),
                        c - 1,
                        checker
                    ),
                    comment,
                    location: toLocation(c, r),
                    ignore: false,
                });
            }
        }

        if (sheet.fields.length > 0) {
            firstSheet ??= sheet;
            workbook.add(sheet);
        }
    }

    if (firstSheet) {
        for (const name in requiredProcessors) {
            if (requiredProcessors[name] === 0) {
                firstSheet.processors.push({
                    name,
                    args: [],
                });
            }
        }
    }
};

export const loadBody = (path: string, data: xlsx.Workbook) => {
    const ctx = getContext(DEFAULT_WRITER, DEFAULT_TAG)!;
    const workbook = ctx.get(path);
    for (const rawSheet of data.getSheets()) {
        if (!workbook.has(rawSheet.name)) {
            continue;
        }
        using _ = trace(`Reading sheet '${rawSheet.name}' in '${path}'`);
        const sheet = workbook.get(rawSheet.name);
        const start = toString(readCell(rawSheet, 1, 1)).startsWith("@")
            ? MAX_HEADERS
            : MAX_HEADERS - 1;
        let maxRows = rawSheet.rowCount;
        for (let r = maxRows; r > start; r--) {
            const cell: TCell | undefined = readCell(rawSheet, r, 1);
            maxRows = r;
            if (cell.v) {
                break;
            }
        }
        const refers: Record<string, { checker: CheckerType; field: Field }> = {};
        for (const field of sheet.fields) {
            for (const checker of field.checkers as CheckerType[]) {
                if (checker.name === BuiltinChecker.Refer) {
                    const name = checker.args[0];
                    const referField = sheet.fields.find((f) => f.name === name);
                    if (!referField) {
                        error(`Refer field not found: ${name} at ${checker.location}`);
                    }
                    referField.ignore = true;
                    refers[name] = { checker, field };
                }
            }
        }
        for (let r = start + 1; r <= maxRows; r++) {
            const row: TRow = {};
            row["!type"] = Type.Row;
            for (const field of sheet.fields) {
                const cell: TCell = readCell(rawSheet, r, field.index + 1);
                if (field.typename === "auto") {
                    if (cell.v !== "-") {
                        error(`Expected '-' at ${toLocation(1, r)}, but got '${cell.v}'`);
                    }
                    cell.v = r - start;
                }
                row[field.name] = cell;
                if (field.index === 0) {
                    sheet.data[r] = row;
                    if (field.name.startsWith("-")) {
                        ignoreField(row, field.name, true);
                        field.ignore = true;
                    }
                } else if (field.typename.startsWith("@")) {
                    const typename = field.typename.slice(1);
                    const refField = sheet.fields.find((f) => f.name === typename);
                    ignoreField(row, typename, true);
                    assert(refField, `Type field not found: ${typename} at ${field.location}`);
                    refField.ignore = true;
                }

                const refer = refers[field.name];
                if (refer && cell.v) {
                    refer.checker.refers[toLocation(refer.field.index + 1, r)] = parseChecker(
                        basename(workbook.path),
                        sheet.name,
                        cell.r,
                        field.index,
                        toString(cell)
                    ) as CheckerType[];
                }
            }
        }
    }
};

export const convertBody = () => {
    const ctx = getContext(DEFAULT_WRITER, DEFAULT_TAG)!;
    for (const workbook of ctx.workbooks) {
        console.log(`parsing: '${workbook.path}'`);
        for (const sheet of workbook.sheets) {
            using _ = trace(`Parsing sheet '${sheet.name}' in '${workbook.path}'`);
            for (const row of Object.values(sheet.data) as TRow[]) {
                if (!row || typeof row !== "object" || row["!type"] !== Type.Row) {
                    continue;
                }
                for (const field of sheet.fields) {
                    const cell = row[field.name];
                    checkType(cell, Type.Cell);
                    let typename = field.typename;
                    if (typename.startsWith("@")) {
                        typename = row[typename.slice(1)]?.v as string;
                        if (!typename) {
                            error(`type not found for ${cell.r}`);
                        }
                    }
                    convertValue(cell, typename);
                }
            }
        }
    }
};
