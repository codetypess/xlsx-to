import * as xlsx from "fastxlsx";
import { basename } from "path";
import { BuiltinChecker, type Checker, type CheckerContext, type CheckerParser, type Convertor, type Processor, type Writer } from "./core/contracts.js";
import { assert, doing, error } from "./core/errors.js";
import {
    addContext,
    clearRunningContext,
    getContext,
    getContexts,
    getRunningContext,
    removeContext,
    setRunningContext,
} from "./core/context-store.js";
import { convertValue, makeCell } from "./core/conversion.js";
import {
    checkerParsers,
    DEFAULT_TAG,
    DEFAULT_WRITER,
    options,
    processors,
    type ProcessorOption,
    type ProcessorType,
    writers,
} from "./core/registry.js";
import { checkType, copyTag, ignoreField, toString } from "./core/value.js";
import { keys, values } from "./util.js";
import { type Field, type Sheet, type TArray, type TCell, type TObject, type TRow, type TValue, Type } from "./core/schema.js";
import { Context, Workbook } from "./core/workbook.js";

export * from "./core/context-store.js";
export * from "./core/conversion.js";
export * from "./core/contracts.js";
export * from "./core/errors.js";
export * from "./core/registry.js";
export * from "./core/schema.js";
export * from "./core/value.js";
export * from "./core/workbook.js";
type CheckerType = {
    readonly name: string;
    readonly force: boolean;
    readonly source: string;
    readonly args: string[];
    readonly location: string;
    readonly refers: Record<string, CheckerType[]>;
    exec: Checker;
};

const MAX_ERRORS = 50;
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
            /**
             * @Processor
             * @Processor(arg1, arg2, ...)
             * @processor({k1,k2}, id, key)
             * @processor([k1,k2], id, key)
             */
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
) => {
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
            using _ = doing(`Parsing checker at ${location}: '${s}'`);
            let checker: CheckerType | undefined;
            if (s.startsWith("@")) {
                /**
                 * @Checker
                 * @Checker(arg1, arg2, ...)
                 */
                const [, name = "", arg = ""] = s.match(/^@(\w+)(?:\((.*?)\))?$/) ?? [];
                checker = {
                    name,
                    force,
                    source: s,
                    location,
                    args: arg.split(",").map((a) => a.trim()),
                    refers: {},
                    exec: null!,
                };
            } else if (s.startsWith("[") && s.endsWith("]")) {
                /**
                 * [0, 1, "a", "b", "c", ...]
                 */
                checker = {
                    name: BuiltinChecker.Range,
                    force,
                    source: s,
                    location,
                    args: [s],
                    refers: {},
                    exec: null!,
                };
            } else if (s.endsWith("#")) {
                /**
                 * file#
                 * #
                 */
                const [, rowKey = "", rowFilter = "", colFile = ""] =
                    s.match(/^(?:\$([^&]*)?(?:&(.+))?==)?([^#]*)#$/) ?? [];
                checker = {
                    name: BuiltinChecker.Sheet,
                    force,
                    source: s,
                    location,
                    args: [rowFile, rowSheet, rowKey, rowFilter, makeFilePath(colFile || rowFile)],
                    refers: {},
                    exec: null!,
                };
            } else if (s.includes("#")) {
                /**
                 * $.id==task#main.id
                 * task#main.id
                 * #main.id
                 * $&key2=MAIN==#main.type&condition=mainline_event
                 */
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
                    refers: {},
                    exec: null!,
                };
            } else if (s !== "x") {
                /**
                 * value >= 0 && value <= 100
                 */
                checker = {
                    name: BuiltinChecker.Expr,
                    force,
                    source: s,
                    location,
                    args: [s],
                    refers: {},
                    exec: null!,
                };
            }
            return checker;
        })
        .filter((v) => !!v);
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

const readHeader = (path: string, data: xlsx.Workbook) => {
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
        using _ = doing(`Reading sheet '${rawSheet.name}' in '${path}'`);
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

const readBody = (path: string, data: xlsx.Workbook) => {
    const ctx = getContext(DEFAULT_WRITER, DEFAULT_TAG)!;
    const workbook = ctx.get(path);
    for (const rawSheet of data.getSheets()) {
        if (!workbook.has(rawSheet.name)) {
            continue;
        }
        using _ = doing(`Reading sheet '${rawSheet.name}' in '${path}'`);
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
            for (const checker of field.checkers) {
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
                    refer.checker.refers[toLocation(refer.field.index, r)] = parseChecker(
                        basename(workbook.path),
                        sheet.name,
                        cell.r,
                        field.index,
                        toString(cell)
                    );
                }
            }
        }
    }
};

const resolveChecker = () => {
    const writerKeys = Object.keys(writers);
    for (const ctx of getContexts()) {
        if (!writerKeys.includes(ctx.writer)) {
            continue;
        }
        for (const workbook of ctx.workbooks) {
            for (const sheet of workbook.sheets) {
                using _ = doing(`Resolving checker in '${workbook.path}#${sheet.name}'`);
                for (const field of sheet.fields) {
                    const checkers = field.checkers.slice();
                    field.checkers.forEach((v) => {
                        if (v.name === BuiltinChecker.Refer) {
                            checkers.push(...Object.values(v.refers).flat());
                        }
                    });
                    for (const checker of checkers) {
                        const parser = checkerParsers[checker.name];
                        if (!parser) {
                            error(
                                `Checker parser not found at ${checker.location}: '${checker.name}'`
                            );
                        }
                        using __ = doing(
                            `Parsing checker at ${checker.location}: ${checker.source}`
                        );
                        assert(!checker.exec, `Checker already parsed: ${checker.location}`);
                        checker.exec = parser(ctx, ...checker.args);
                    }
                }
            }
        }
    }
};

const parseBody = () => {
    const ctx = getContext(DEFAULT_WRITER, DEFAULT_TAG)!;
    for (const workbook of ctx.workbooks) {
        console.log(`parsing: '${workbook.path}'`);
        for (const sheet of workbook.sheets) {
            using _ = doing(`Parsing sheet '${sheet.name}' in '${workbook.path}'`);
            for (const row of values<TRow>(sheet.data)) {
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

const copyWorkbook = () => {
    for (const ctx of getContexts().slice()) {
        for (const writer in writers) {
            if (options.suppressWriters.includes(writer)) {
                continue;
            }
            console.log(`creating context: writer=${writer} tag=${ctx.tag}`);
            const newCtx = addContext(new Context(writer, ctx.tag));
            for (const workbook of ctx.workbooks) {
                for (const sheet of workbook.sheets) {
                    using _ = doing(`Checking sheet '${sheet.name}' in '${workbook.path}'`);
                    const data: TObject = {};
                    copyTag(sheet.data, data);
                    const keyField = sheet.fields[0];
                    for (const row of values<TRow>(sheet.data)) {
                        const key = row[keyField.name].v as string;
                        if (key === "" || key === undefined || key === null) {
                            error(`Key is empty at ${row[keyField.name].r}`);
                        }
                        if (data[key]) {
                            const last = (data[key] as TRow)[keyField.name];
                            const curr = row[keyField.name];
                            error(`Duplicate key: ${key}, last: ${last.r}, current: ${curr.r}`);
                        }
                        data[key] = row;
                    }
                    sheet.data = data;
                }
                newCtx.add(workbook.clone(newCtx));
            }
        }
    }
};

const invokeReferChecker = (
    ctx: CheckerContext,
    cell: TCell,
    checkers: CheckerType[],
    errors: string[]
) => {
    for (const checker of checkers) {
        const errorValues: string[] = [];
        if ((cell.v !== null || checker.force) && !checker.exec(ctx)) {
            errorValues.push(`${cell.r}: ${cell.s}`);
            if (ctx.errors.length > 0) {
                for (const str of ctx.errors) {
                    errorValues.push("    ❌ " + str);
                }
                ctx.errors.length = 0;
            }
        }
        if (errorValues.length > 0) {
            if (errorValues.length > MAX_ERRORS) {
                errorValues.length = MAX_ERRORS;
                errorValues.push("...");
            }
            errors.push(
                `builtin check:\n` +
                    `     path: ${ctx.workbook.path}\n` +
                    `    sheet: ${ctx.sheet.name}\n` +
                    `    field: ${ctx.field.name}\n` +
                    `  checker: ${checker.source}\n` +
                    `   values:\n` +
                    `      ${errorValues.join("\n      ")}\n`
            );
        }
    }
};

const invokeChecker = (workbook: Workbook, sheet: Sheet, field: Field, errors: string[]) => {
    const checkers = field.checkers.filter((c) => !options.suppressCheckers.includes(c.name));
    const ctx: CheckerContext = {
        workbook,
        sheet,
        field,
        errors: [],
        cell: null!,
        row: null!,
    };
    for (const checker of checkers) {
        const errorValues: string[] = [];
        for (const row of values<TRow>(sheet.data)) {
            const cell = row[field.name];
            checkType(cell, Type.Cell);
            ctx.cell = cell;
            ctx.row = row;
            if ((cell.v !== null || checker.force) && !checker.exec(ctx)) {
                errorValues.push(`${cell.r}: ${cell.s}`);
                if (ctx.errors.length > 0) {
                    for (const str of ctx.errors) {
                        errorValues.push("    ❌ " + str);
                    }
                    ctx.errors.length = 0;
                }
            }
            if (checker.name === BuiltinChecker.Refer) {
                const refers = checker.refers[cell.r];
                if (refers) {
                    invokeReferChecker(ctx, cell, refers, errors);
                }
            }
        }
        if (errorValues.length > 0) {
            if (errorValues.length > MAX_ERRORS) {
                errorValues.length = MAX_ERRORS;
                errorValues.push("...");
            }
            errors.push(
                `builtin check:\n` +
                    `     path: ${workbook.path}\n` +
                    `    sheet: ${sheet.name}\n` +
                    `    field: ${field.name}\n` +
                    `  checker: ${checker.source}\n` +
                    `   values:\n` +
                    `      ${errorValues.join("\n      ")}\n`
            );
        }
    }
};

const performChecker = () => {
    const writerKeys = Object.keys(writers);
    for (const ctx of getContexts()) {
        if (!writerKeys.includes(ctx.writer)) {
            continue;
        }
        console.log(`performing checker: writer=${ctx.writer} tag=${ctx.tag}`);
        const errors: string[] = [];
        for (const workbook of ctx.workbooks) {
            for (const sheet of workbook.sheets) {
                for (const field of sheet.fields) {
                    const msg = `'${field.name}' at ${field.location} in '${workbook.path}#${sheet.name}'`;
                    using _ = doing(`Checking ${msg}`);
                    try {
                        invokeChecker(workbook, sheet, field, errors);
                    } catch (e) {
                        error((e as Error).stack ?? String(e));
                    }
                }
            }
        }
        if (errors.length > 0) {
            throw new Error(`tag: ${ctx.tag} writer: ${ctx.writer}\n` + errors.join("\n"));
        }
    }
};

const performProcessor = async (stage: ProcessorOption["stage"], writer?: string) => {
    type ProcessorEntry = {
        processor: ProcessorType;
        sheet: Sheet;
        args: string[];
        name: string;
    };
    const writerKeys = writer ? [writer] : Object.keys(writers);
    for (const ctx of getContexts().slice()) {
        if (!writerKeys.includes(ctx.writer)) {
            continue;
        }
        setRunningContext(ctx);
        console.log(`performing processor: stage=${stage} writer=${ctx.writer} tag=${ctx.tag}`);
        for (const workbook of ctx.workbooks) {
            const arr: ProcessorEntry[] = [];
            for (const sheet of workbook.sheets) {
                for (const { name, args } of sheet.processors) {
                    const processor = processors[name];
                    if (
                        processor.option.stage !== stage ||
                        options.suppressProcessors.includes(name)
                    ) {
                        continue;
                    }
                    arr.push({
                        processor: processor,
                        sheet: sheet,
                        args: args,
                        name: name,
                    });
                }
            }
            arr.sort((a, b) => a.processor.option.priority - b.processor.option.priority);
            for (const { processor, sheet, args, name } of arr) {
                using _ = doing(
                    `Performing processor '${name}' in '${workbook.path}#${sheet.name}'`
                );
                try {
                    await processor.exec(workbook, sheet, ...args);
                } catch (e) {
                    error((e as Error).stack ?? String(e));
                }
            }
        }
        clearRunningContext();
    }
};

export const parse = async (fs: string[], headerOnly: boolean = false) => {
    const ctx = addContext(new Context(DEFAULT_WRITER, DEFAULT_TAG));
    for (const file of fs) {
        ctx.add(new Workbook(ctx, file));
    }
    for (const file of fs) {
        console.log(`reading: '${file}'`);
        const data = await xlsx.Workbook.open(file);
        readHeader(file, data);
        if (!headerOnly) {
            readBody(file, data);
        }
    }
    await performProcessor("after-read", DEFAULT_WRITER);
    if (!headerOnly) {
        await performProcessor("pre-parse", DEFAULT_WRITER);
        parseBody();
        await performProcessor("after-parse", DEFAULT_WRITER);
        copyWorkbook();
        await performProcessor("pre-check");
        resolveChecker();
        performChecker();
        await performProcessor("after-check");
        await performProcessor("pre-stringify");
        await performProcessor("stringify");
        await performProcessor("after-stringify");
    }
};

export const write = (workbook: Workbook, processor: string, data: object) => {
    const writer = workbook.context.writer;
    assert(!!writers[writer], `Writer not found: ${writer}`);
    writers[writer](workbook, processor, data as TObject | TArray);
};
