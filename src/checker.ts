import { ColumnIndexer, RowFilter } from "./indexer";
import { keys } from "./util";
import { CheckerParser, convertValue, error, getWorkbook, TCell, TObject, TValue } from "./xlsx";

export const SizeCheckerParser: CheckerParser = (arg) => {
    const length = Number(arg);
    if (isNaN(length)) {
        throw new Error(`Invalid length: '${length}'`);
    }
    return (cell, row, field, errors) => {
        if (cell.v instanceof Array) {
            return cell.v.length === length;
        }
        return false;
    };
};

export const ExprCheckerParser: CheckerParser = (arg) => {
    const expr = new Function("$", "return " + arg);
    return (cell, row, field, errors) => {
        return expr(cell.v);
    };
};

export const FollowCheckerParser: CheckerParser = (arg) => {
    return (cell, row, field, errors) => {
        const follow = row[arg] as TCell;
        if (follow.v !== null) {
            return cell.v !== null;
        } else {
            return cell.v === null;
        }
    };
};

export const RangeCheckerParser: CheckerParser = (arg) => {
    let values: unknown[] = [];
    try {
        values = JSON.parse(arg);
    } catch (e) {
        throw new Error(`Invalid range: '${arg}'`);
    }
    return (cell, row, field, errors) => {
        return values.includes(cell.v);
    };
};

const parseResolver = (expr: IndexerFilterExpr) => {
    type Collector = (value: TValue, collector: TValue[]) => void;
    const collectors: Collector[] = [];
    let str = expr.key.trim().replaceAll(" ", "");

    while (str.length) {
        const [match, query, optional] = str.match(/^(\.\w+|\[\d+\]|\[\*\]|\[\.\])([?]?)/) ?? [];
        if (match) {
            str = str.slice(match.length);
            if (query.startsWith(".")) {
                const key = query.slice(1);
                collectors.push((value, arr) => {
                    if (value && typeof value === "object") {
                        const v = (value as TObject)[key];
                        if (v !== undefined || !optional) {
                            arr.push(v);
                        }
                    } else {
                        arr.push(null);
                    }
                });
            } else if (query === "[*]") {
                collectors.push((value, arr) => {
                    if (Array.isArray(value)) {
                        for (const item of value) {
                            arr.push(item);
                        }
                    } else {
                        arr.push(null);
                    }
                });
            } else if (query === "[.]") {
                collectors.push((value, arr) => {
                    if (value && typeof value === "object") {
                        arr.push(...keys(value as TObject));
                    } else {
                        arr.push(null);
                    }
                });
            } else {
                const index = Number(query.slice(1, -1));
                collectors.push((value, arr) => {
                    if (Array.isArray(value)) {
                        const v = value[index];
                        if (v !== undefined || !optional) {
                            arr.push(v);
                        }
                    } else {
                        arr.push(null);
                    }
                });
            }
        } else {
            throw new Error(`Invalid query: ${expr}`);
        }
    }

    const arr: TValue[] = [];
    return (value: TValue, errors: string[], walker: (value: string | number) => boolean) => {
        arr.length = 0;
        arr.push(value);
        let start = 0;
        for (const query of collectors) {
            const length = arr.length;
            for (let i = start; i < length; i++) {
                query(arr[i], arr);
            }
            start = length;
        }
        for (let i = start; i < arr.length; i++) {
            const v = arr[i];
            if (!(typeof v === "string" || typeof v === "number") || !walker(v)) {
                errors.push(`${v}`);
                return false;
            }
        }
        return true;
    };
};

type IndexerFilterExpr = {
    file: string;
    sheet: string;
    key: string;
    filter: string;
};

const parseFilter = (expr: IndexerFilterExpr) => {
    const workbook = getWorkbook(expr.file);
    const findField = (name: string) => {
        if (expr.sheet === "*") {
            for (const sheet of Object.values(workbook.sheets)) {
                const field = sheet.fields.find((f) => f.name === name);
                if (field) {
                    return field;
                }
            }
        } else {
            const sheet = workbook.sheets[expr.sheet];
            return sheet.fields.find((f) => f.name === name);
        }
    };
    return expr.filter
        .replaceAll(" ", "")
        .split("&")
        .filter((s) => s.length)
        .map((s) => {
            const [, key, value] = s.match(/(\w+)=(\w+)/) ?? [];
            if (key && value) {
                const field = findField(key);
                if (!field) {
                    error(`Field not found: ${key}`);
                }
                const v = convertValue(value, field.typename);
                return { key, value: v };
            } else {
                error(`Invalid filter: ${expr.filter}`);
            }
        }) as readonly RowFilter[];
};

const parseIndexerAst = (rowExpr: IndexerFilterExpr, colExpr: IndexerFilterExpr) => {
    return {
        value: {
            key: rowExpr.key,
            resolve: parseResolver(rowExpr),
            filter: parseFilter(rowExpr),
        },
        target: {
            key: colExpr.key,
            filter: parseFilter(colExpr),
        },
    };
};

export const IndexCheckerParser: CheckerParser = (
    rowFile,
    rowSheet,
    rowKey,
    rowFilter,
    colFile,
    colSheet,
    colKey,
    colFilter
) => {
    const ast = parseIndexerAst(
        {
            file: rowFile,
            sheet: rowSheet,
            key: rowKey,
            filter: rowFilter,
        },
        {
            file: colFile,
            sheet: colSheet,
            key: colKey,
            filter: colFilter,
        }
    );
    const indexer = new ColumnIndexer(colFile, colSheet, ast.target.key);

    return (cell, row, field, errors) => {
        if (cell.v === null || cell.v === undefined) {
            throw new Error(`Invalid value at ${cell.r} in ${field.path}#${field.sheet}`);
        }

        if (ast.value.filter.length > 0) {
            // skip cell if not match any filter
            for (const entry of ast.value.filter) {
                const rowCell = row[entry.key] as TCell | undefined;
                if (!rowCell) {
                    throw new Error(
                        `field '${entry.key}' not found in ${field.path}#${field.sheet}`
                    );
                }
                if (rowCell.v !== entry.value) {
                    return true;
                }
            }
        }

        return ast.value.resolve(cell.v, errors, (value) => {
            if (ast.target.filter.length) {
                return indexer.has(value, ast.target.filter);
            } else {
                return indexer.has(value);
            }
        });
    };
};

export const SheetCheckerParser: CheckerParser = (file) => {
    const path = file.replace(/\.xlsx$/, "") + ".xlsx";
    const workbook = getWorkbook(path);
    return (cell, row, field, errors) => {
        const sheet = workbook.sheets[cell.v as string];
        return sheet !== undefined;
    };
};
