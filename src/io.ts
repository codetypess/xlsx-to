import * as fs from "node:fs";
import { dirname } from "node:path";
import { assert } from "./core/errors.js";
import { writers } from "./core/registry.js";
import type { TArray, TObject, TValue } from "./core/schema.js";
import type { Workbook } from "./core/workbook.js";
import {
    JsonStringifyOption,
    LuaStringifyOption,
    stringifyJson,
    stringifyLua,
    stringifyTs,
    TsStringifyOption,
} from "./stringify.js";

export const readFile = (path: string) => {
    if (fs.existsSync(path)) {
        return fs.readFileSync(path, "utf-8");
    }
    return null;
};

export const writeFile = (path: string, data: string) => {
    const dir = dirname(path);
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
    if (fs.existsSync(path) && readFile(path) === data) {
        console.log(`up-to-date: ${path}`);
    } else {
        console.log(`write: ${path}`);
        fs.writeFileSync(path, data, { encoding: "utf-8" });
    }
};

export const writeJson = (path: string, data: unknown, options?: JsonStringifyOption) => {
    options = options ?? {};
    options.indent = options.indent ?? 2;
    options.precision = options.precision ?? 10;
    writeFile(path, stringifyJson(data as TValue, options));
};

export const writeLua = (path: string, data: unknown, options?: LuaStringifyOption) => {
    options = options ?? {};
    options.indent = options.indent ?? 2;
    options.precision = options.precision ?? 10;
    writeFile(path, stringifyLua(data as TValue, options));
};

export const writeTs = (path: string, data: unknown, options?: TsStringifyOption) => {
    options = options ?? {};
    options.indent = options.indent ?? 2;
    options.precision = options.precision ?? 10;
    writeFile(path, stringifyTs(data as TValue, options));
};

export const write = (workbook: Workbook, processor: string, data: object) => {
    const writer = workbook.context.writer;
    assert(!!writers[writer], `Writer not found: ${writer}`);
    writers[writer](workbook, processor, data as TObject | TArray);
};
