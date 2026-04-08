import JSON5 from "json5";
import { Converter } from "../core/contracts.js";

const bools: Record<string, boolean> = {
    ["true"]: true,
    ["1"]: true,
    ["✔︎"]: true,
    ["false"]: false,
    ["0"]: false,
    ["✖︎"]: false,
    ["x"]: false,
};

export const boolConverter: Converter = (str) => {
    return bools[str] ?? null;
};

export const intConverter: Converter = (str) => {
    if (str === "") {
        return null;
    }
    const result = Number(str);
    if (isNaN(result) || result !== (result | 0)) {
        return null;
    }
    return result;
};

export const stringConverter: Converter = (str) => {
    return str === "" ? null : str;
};

export const floatConverter: Converter = (str) => {
    if (str === "") {
        return null;
    }
    const result = Number(str);
    if (isNaN(result)) {
        return null;
    }
    return result;
};

export const jsonConverter: Converter = (str) => {
    try {
        return JSON5.parse(str);
        // eslint-disable-next-line no-empty
    } catch (_) {}
    try {
        return JSON.parse(str);
        // eslint-disable-next-line no-empty
    } catch (_) {}
    return null;
};
