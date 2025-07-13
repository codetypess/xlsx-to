import { escape, filterKeys, isNullOrUndefined, isNumericKey } from "./util";
import { TArray, TCell, TObject, TValue, Type } from "./xlsx";

export class StringBuffer {
    readonly data: string[] = [];

    private _indent: number;
    private _indentCount: number = 0;

    constructor(indent: number) {
        this._indent = indent;
    }

    get indentCount() {
        return this._indentCount;
    }

    indent() {
        this._indentCount += this._indent;
    }

    unindent() {
        this._indentCount -= this._indent;
    }

    padding() {
        if (this._indent > 0) {
            this.data.push(" ".repeat(this._indentCount));
        }
    }

    linefeed() {
        if (this._indent > 0) {
            this.data.push("\n");
        }
    }

    writeLine(value: string) {
        this.padding();
        this.data.push(value);
        this.linefeed();
    }

    writeLines(value: string) {
        for (const line of value.split("\n")) {
            this.writeLine(line);
        }
    }

    writeString(value: string) {
        this.data.push(value);
    }

    toString() {
        return this.data.join("");
    }
}

//-----------------------------------------------------------------------------
// Json
//-----------------------------------------------------------------------------
type JsonStringifyOption = {
    indent?: number;
    precision?: number;
};

export const stringifyJson = (data: TValue, option?: JsonStringifyOption) => {
    const stacks: string[] = [];
    option = option ?? {};
    option.indent = Math.max(option.indent ?? 4, 0);
    const buffer = new StringBuffer(option.indent);

    function writeJsonValue(value: TValue, asKeyValue: boolean) {
        if (!asKeyValue) {
            buffer.padding();
        }
        if (typeof value === "number") {
            if (value === (value | 0)) {
                buffer.writeString(value.toFixed(0));
            } else {
                buffer.writeString(value.toFixed(option?.precision).replace(/\.?0+$/, ""));
            }
        } else if (typeof value === "boolean") {
            buffer.writeString(value.toString());
        } else if (value === null || value === undefined) {
            buffer.writeString("null");
        } else if (typeof value === "string") {
            buffer.writeString('"');
            buffer.writeString(escape(value));
            buffer.writeString('"');
        } else if (Array.isArray(value)) {
            writeJsonArray(value);
        } else {
            let item: TValue = value;
            if (item["!type"] === Type.Cell) {
                const cell = item as unknown as TCell;
                item = cell.v as TValue;
            }
            if (typeof item !== "object" || item === null || item === undefined) {
                writeJsonValue(item, true);
            } else if (Array.isArray(item)) {
                writeJsonArray(item);
            } else {
                writeJsonObject(item as TObject);
            }
        }
    }

    function writeJsonObject(value: TObject) {
        if (stacks.length > 256) {
            throw new Error(`json stringify stack overflow: ${stacks.join("->")}`);
        }

        if (value["!toString"]) {
            buffer.writeLines(value["!toString"](value, option!.indent!, "json"));
            return;
        }

        const keys = filterKeys(value, true, (v) => !isNullOrUndefined(v));
        const space = option!.indent! > 0 ? " " : "";
        buffer.writeString("{");
        buffer.linefeed();
        buffer.indent();
        for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            const v = value[k];
            stacks.push(k);
            buffer.padding();
            buffer.writeString(`"${k}":${space}`);
            writeJsonValue(v, true);
            if (i < keys.length - 1) {
                buffer.writeString(",");
            }
            buffer.linefeed();
            stacks.pop();
        }
        buffer.unindent();
        buffer.padding();
        buffer.writeString("}");
    }

    function writeJsonArray(value: TArray) {
        if (value["!toString"]) {
            buffer.writeLines(value["!toString"](value, option!.indent!, "json"));
            return;
        }

        buffer.writeString("[");
        buffer.linefeed();
        buffer.indent();
        for (let i = 0; i < value.length; i++) {
            const v = value[i];
            writeJsonValue(v, false);
            if (i < value.length - 1) {
                buffer.writeString(",");
            }
            buffer.linefeed();
        }
        buffer.unindent();
        buffer.padding();
        buffer.writeString("]");
    }

    writeJsonValue(data, false);

    return buffer.toString();
};

//-----------------------------------------------------------------------------
// Lua
//-----------------------------------------------------------------------------
type LuaStringifyOption = {
    indent?: number;
    marshal?: string;
    precision?: number;
};

export const stringifyLua = (data: TValue, option?: LuaStringifyOption) => {
    const stacks: string[] = [];
    option = option ?? {};
    option.indent = Math.max(option.indent ?? 4, 0);
    const buffer = new StringBuffer(option.indent);

    function writeLuaComment(value?: TValue) {
        if (value && typeof value === "object") {
            if (value["!enum"]) {
                buffer.writeLine(`---@enum ${value["!enum"]}`);
            }
            if (value["!comment"]) {
                const comment = value["!comment"].replaceAll("\n", "\\n");
                buffer.writeLine(`-- ${comment}`);
            }
        }
    }

    function writeLuaValue(value: TValue, asKeyValue: boolean) {
        if (!asKeyValue) {
            buffer.padding();
        }
        if (typeof value === "number") {
            if (value === (value | 0)) {
                buffer.writeString(value.toFixed(0));
            } else {
                buffer.writeString(value.toFixed(option?.precision).replace(/\.?0+$/, ""));
            }
        } else if (typeof value === "boolean") {
            buffer.writeString(value.toString());
        } else if (value === null || value === undefined) {
            buffer.writeString("nil");
        } else if (typeof value === "string") {
            buffer.writeString('"');
            buffer.writeString(escape(value));
            buffer.writeString('"');
        } else if (Array.isArray(value)) {
            writeLuaArray(value);
        } else {
            let item: TValue = value;
            if (item["!type"] === Type.Cell) {
                const cell = item as unknown as TCell;
                item = cell.v as TValue;
            }
            if (typeof item !== "object" || item === null || item === undefined) {
                writeLuaValue(item, true);
            } else if (Array.isArray(item)) {
                writeLuaArray(item);
            } else {
                writeLuaObject(item as TObject);
            }
        }
    }

    function writeLuaObject(value: TObject) {
        if (stacks.length > 256) {
            throw new Error(`json stringify stack overflow: ${stacks.join("->")}`);
        }

        if (value["!toString"]) {
            buffer.writeLines(value["!toString"](value, option!.indent!, "lua"));
            return;
        }

        const keys = filterKeys(value, true, (v) => !isNullOrUndefined(v));
        const space = option!.indent! > 0 ? " " : "";
        buffer.writeString("{");
        buffer.linefeed();
        buffer.indent();
        for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            const v = value[k];
            stacks.push(k);
            writeLuaComment(v);
            buffer.padding();
            if (isNumericKey(k)) {
                buffer.writeString(`[${k}]${space}=${space}`);
            } else if (k.match(/^[a-zA-Z_][a-zA-Z0-9_]*$/)) {
                buffer.writeString(`${k}${space}=${space}`);
            } else {
                buffer.writeString(`["${k}"]${space}=${space}`);
            }
            writeLuaValue(v, true);
            buffer.writeString(",");
            buffer.linefeed();
            stacks.pop();
        }
        buffer.unindent();
        buffer.padding();
        buffer.writeString("}");
    }

    function writeLuaArray(value: TArray) {
        if (value["!toString"]) {
            buffer.writeLines(value["!toString"](value, option!.indent!, "lua"));
            return;
        }

        buffer.writeString("{");
        buffer.linefeed();
        buffer.indent();
        for (let i = 0; i < value.length; i++) {
            const v = value[i];
            writeLuaValue(v, false);
            buffer.writeString(",");
            buffer.linefeed();
        }
        buffer.unindent();
        buffer.padding();
        buffer.writeString("}");
    }

    if (option.marshal) {
        buffer.writeString(option.marshal);
    }

    writeLuaValue(data, false);

    return buffer.toString();
};

//-----------------------------------------------------------------------------
// TypeScript
//-----------------------------------------------------------------------------
type TsStringifyOption = {
    indent?: number;
    marshal?: string;
    precision?: number;
};

export const stringifyTs = (data: TValue, option?: TsStringifyOption) => {
    const stacks: string[] = [];
    option = option ?? {};
    option.indent = Math.max(option.indent ?? 4, 0);
    const buffer = new StringBuffer(option.indent);
    const enumBuffer = new StringBuffer(option.indent);

    function writeTsComment(comment: string | undefined, out: StringBuffer) {
        if (comment) {
            out.writeLine(`/**`);
            comment.split("\n").forEach((line) => {
                out.writeLine(` * ${line}`);
            });
            out.writeLine(` */`);
        }
    }

    function writeTsValue(value: TValue, asKeyValue: boolean) {
        if (!asKeyValue) {
            buffer.padding();
        }
        if (typeof value === "number") {
            if (value === (value | 0)) {
                buffer.writeString(value.toFixed(0));
            } else {
                buffer.writeString(value.toFixed(option?.precision).replace(/\.?0+$/, ""));
            }
        } else if (typeof value === "boolean") {
            buffer.writeString(value.toString());
        } else if (value === null || value === undefined) {
            buffer.writeString("null");
        } else if (typeof value === "string") {
            buffer.writeString('"');
            buffer.writeString(escape(value));
            buffer.writeString('"');
        } else if (Array.isArray(value)) {
            writeTsArray(value);
        } else if (value["!enum"]) {
            const enumName = value["!enum"];
            const enumComment = value["!comment"];
            const keys = filterKeys(value as TObject, true, (v) => !isNullOrUndefined(v));
            writeTsComment(enumComment, enumBuffer);
            enumBuffer.writeLine(`export enum ${enumName} {`);
            enumBuffer.indent();
            for (const k of keys) {
                const v = (value as TObject)[k] as TCell;
                const valueComment = v["!comment"];
                writeTsComment(valueComment, enumBuffer);
                if (typeof v.v === "number") {
                    enumBuffer.writeLine(`${k} = ${v.v},`);
                } else {
                    enumBuffer.writeLine(`${k} = "${v.v}",`);
                }
            }
            enumBuffer.unindent();
            enumBuffer.writeLine(`}`);
            enumBuffer.writeLine("");
            buffer.writeString(enumName);
        } else {
            let item: TValue = value;
            if (item["!type"] === Type.Cell) {
                const cell = item as unknown as TCell;
                item = cell.v as TValue;
            }
            if (typeof item !== "object" || item === null || item === undefined) {
                writeTsValue(item, true);
            } else if (Array.isArray(item)) {
                writeTsArray(item);
            } else {
                writeTsObject(item as TObject);
            }
        }
    }

    function writeTsObject(value: TObject) {
        if (stacks.length > 256) {
            throw new Error(`json stringify stack overflow: ${stacks.join("->")}`);
        }

        if (value["!toString"]) {
            buffer.writeLines(value["!toString"](value, option!.indent!, "ts"));
            return;
        }

        const keys = filterKeys(value, true, (v) => !isNullOrUndefined(v));
        const space = option!.indent! > 0 ? " " : "";
        buffer.writeString("{");
        buffer.linefeed();
        buffer.indent();
        for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            const v = value[k];
            stacks.push(k);
            if (v && typeof v === "object" && v["!comment"]) {
                writeTsComment(v["!comment"], buffer);
            }
            buffer.padding();
            if (k.match(/^[a-zA-Z_$][a-zA-Z0-9_$]*$/) || isNumericKey(k)) {
                buffer.writeString(`${k}:${space}`);
            } else {
                buffer.writeString(`"${k}":${space}`);
            }
            writeTsValue(v, true);
            buffer.writeString(",");
            buffer.linefeed();
            stacks.pop();
        }
        buffer.unindent();
        buffer.padding();
        buffer.writeString("}");
    }

    function writeTsArray(value: TArray) {
        if (value["!toString"]) {
            buffer.writeLines(value["!toString"](value, option!.indent!, "ts"));
            return;
        }

        buffer.writeString("[");
        buffer.linefeed();
        buffer.indent();
        for (let i = 0; i < value.length; i++) {
            const v = value[i];
            writeTsValue(v, false);
            buffer.writeString(",");
            buffer.linefeed();
        }
        buffer.unindent();
        buffer.padding();
        buffer.writeString("]");
    }

    if (option.marshal) {
        buffer.writeString(option.marshal);
    }

    writeTsValue(data, false);

    const enumString = enumBuffer.toString();
    if (enumString) {
        buffer.data.unshift("\n");
        buffer.data.unshift(enumString);
    }

    return buffer.toString();
};
