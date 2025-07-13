import { CheckerParser, createColumnIndexer, error, TCell } from "./xlsx";

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

export const IndexCheckerParser: CheckerParser = (file, sheetName, key, idx) => {
    const indexer = createColumnIndexer(file, sheetName, key);

    return (cell, row, field, errors) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const value: any = cell.v;
        if (value === null || value === undefined) {
            error(`Invalid value at ${cell.r} in ${field.path}#${field.sheet}`);
        }
        if (typeof value !== "object") {
            return indexer.has(value);
        } else if (Array.isArray(value)) {
            /**
             * [value, value, ...]
             * [{idx: value}, {idx: value}, ...]
             */
            let found = 0;
            for (const item of value) {
                if (!idx) {
                    if (indexer.has(item)) {
                        found++;
                    }
                } else if (typeof item === "object" && indexer.has(item[idx])) {
                    found++;
                }
            }
            return found === value.length;
        } else {
            /**
             * {idx: value}
             */
            return typeof value === "object" && indexer.has(value[idx]);
        }
    };
};
