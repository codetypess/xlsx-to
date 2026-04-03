import { Type, type Tag, type TCell, type TObject, type TValue } from "./schema.js";

export const copyTag = (src: object & Tag, dest: object & Tag) => {
    Object.keys(src)
        .filter((k) => k.startsWith("!"))
        .forEach((k) => ((dest as TObject)[k] = (src as TObject)[k]));
};

export const typeOf = (value: TValue) => {
    if (value && typeof value === "object" && value["!type"]) {
        return value["!type"];
    }
    return typeof value;
};

export const checkType = <T>(value: TValue, type: Type | string) => {
    const t = typeOf(value);
    if (t === type) {
        return value as T;
    }
    console.error(`checking value: `, value);
    throw new Error(`Expect type '${type}', but got '${t}'`);
};

export const isNull = (value: TValue): value is null | undefined => {
    if (value === null || value === undefined) {
        return true;
    }
    if (typeof value === "object" && value["!type"] === Type.Cell) {
        const cell = value as unknown as TCell;
        if (cell.v === null || cell.v === undefined) {
            return true;
        }
    }
    return false;
};

export const isNotNull = (value: TValue): value is Exclude<TValue, null | undefined> => {
    return !isNull(value);
};

export const ignoreField = (obj: object & Tag, field: string, ignored: boolean) => {
    obj["!ignore"] ??= {};
    obj["!ignore"][field] = ignored;
};

export const toString = (cell?: TCell) => {
    if (isNull(cell)) {
        return "";
    }
    if (typeof cell.v === "string") {
        return cell.v.trim();
    }
    return String(cell.v);
};
