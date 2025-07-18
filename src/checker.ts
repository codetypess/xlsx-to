import { ColumnIndexer, RowFilter } from "./indexer";
import { assert, CheckerParser, error, get, TCell } from "./xlsx";

export const SizeCheckerParser: CheckerParser = (arg) => {
    const length = Number(arg);
    if (isNaN(length)) {
        error(`Invalid length: '${length}'`);
    }
    return (cell, row, field, errors) => {
        if (cell.v instanceof Array) {
            return cell.v.length === length;
        }
        return false;
    };
};

export const ExprCheckerParser: CheckerParser = (arg) => {
    const expr = new Function("value", "return " + arg);
    return (cell, row, field, errors) => {
        try {
            return expr(cell.v);
        } catch (e) {
            error(String(e));
        }
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
        error(`Invalid range: '${arg}'`);
    }
    return (cell, row, field, errors) => {
        return values.includes(cell.v);
    };
};

type ValueQuery = {
    readonly key: string;
    readonly filter: readonly RowFilter[];
};

const parseQuery = (query?: string): ValueQuery => {
    if (!query) {
        return { key: "", filter: [] };
    }
    const mainKey = query.match(/^\w+/)?.[0] ?? "";
    const filter = Array.from(query.matchAll(/([^=&]+)=([^=&]+)/g)).map(([_, key, value]) => {
        key = key.trim();
        value = value.trim();
        const num = Number(value);
        return { key: key, value: isNaN(num) ? value : num };
    });
    return { key: mainKey, filter };
};

export const IndexCheckerParser: CheckerParser = (file, sheetName, rowQuery, columnQuery) => {
    const queryColumn = parseQuery(columnQuery);
    const queryRow = parseQuery(rowQuery);

    assert(!!queryColumn.key, `Invalid key: '${columnQuery}'`);

    const indexer = new ColumnIndexer(file, sheetName, queryColumn.key);

    const check = (value: unknown) => {
        if (queryColumn.filter.length) {
            return indexer.has(value as string | number, queryColumn.filter);
        } else {
            return indexer.has(value as string | number);
        }
    };

    return (cell, row, field, errors) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const value: any = cell.v;
        if (value === null || value === undefined) {
            error(`Invalid value at ${cell.r} in ${field.path}#${field.sheet}`);
        }

        if (queryRow.filter.length > 0) {
            // skip cell if not match any filter
            for (const { key, value } of queryRow.filter) {
                const cell = row[key] as TCell | undefined;
                if (!cell || cell.v !== value) {
                    return true;
                }
            }
        }

        if (typeof value !== "object") {
            return check(value);
        } else if (Array.isArray(value)) {
            /**
             * [value, value, ...]
             * [{idx: value}, {idx: value}, ...]
             */
            let found = 0;
            for (const item of value) {
                if (!queryRow.key) {
                    if (check(item)) {
                        found++;
                    }
                } else if (typeof item === "object" && check(item[queryRow.key])) {
                    found++;
                }
            }
            return found === value.length;
        } else {
            /**
             * {idx: value}
             */
            return typeof value === "object" && check(value[queryRow.key]);
        }
    };
};

export const SheetCheckerParser: CheckerParser = (file) => {
    return (cell, row, field, errors) => {
        const path = file.replace(/\.xlsx$/, "") + ".xlsx";
        const workbook = get(path);
        const sheet = workbook.sheets[cell.v as string];
        return sheet !== undefined;
    };
};
