import { StringBuffer } from "./stringify";
import { filename, toPascalCase } from "./util";
import { convertors, files, get } from "./xlsx";

type ClassNameMaker = (className: string) => string;

export const genTsTypedef = (path: string, writer: string, maker?: ClassNameMaker) => {
    const workbook = get(path);
    const buffer = new StringBuffer(4);
    const name = filename(path);
    maker = maker ?? ((className) => className);
    for (const k of Object.keys(workbook.sheets).sort()) {
        const sheet = workbook.sheets[k];
        const className = maker(toPascalCase(`Generated_${name}_${sheet.name}_Row`));

        // row
        buffer.writeLine(`// file: ${path}`);
        buffer.writeLine(`export interface ${className} {`);
        buffer.indent();
        for (const field of sheet.fields) {
            if (!field.writers.includes(writer)) {
                continue;
            }
            const checker = field.checker.map((v) => v.def).join(";");
            const optional = field.typename.endsWith("?") ? "?" : "";
            const comment = field.comment.replaceAll(/[\r\n]+/g, " ");
            let typename = field.typename.replaceAll("?", "");
            typename = convertors[typename].realtype ?? typename;
            buffer.writeLine(`/**`);
            buffer.writeLine(
                ` * ${comment} (location: ${field.refer}) (checker: ${checker || "x"})`
            );
            buffer.writeLine(` */`);
            if (typename === "int" || typename === "float") {
                buffer.writeLine(`readonly ${field.name}${optional}: number;`);
            } else if (typename === "string") {
                buffer.writeLine(`readonly ${field.name}${optional}: string;`);
            } else if (typename === "bool") {
                buffer.writeLine(`readonly ${field.name}${optional}: boolean;`);
            } else {
                buffer.writeLine(`readonly ${field.name}${optional}: unknown;`);
            }
        }
        buffer.unindent();
        buffer.writeLine(`}`);
        buffer.writeLine("");

        // col
        buffer.writeLine(`// file: ${path}`);
        buffer.writeLine(`export interface ${className.replace("Row", "Col")} {`);
        buffer.indent();
        for (const field of sheet.fields) {
            if (!field.writers.includes(writer)) {
                continue;
            }
            const checker = field.checker.map((v) => v.def).join(";");
            const comment = field.comment.replaceAll(/[\r\n]+/g, " ");
            const optional = field.typename.endsWith("?") ? " | undefined" : "";
            let typename = field.typename.replaceAll("?", "");
            typename = convertors[typename].realtype ?? typename;
            buffer.writeLine(`/**`);
            buffer.writeLine(
                ` * ${comment} (location: ${field.refer}) (checker: ${checker || "x"})`
            );
            buffer.writeLine(` */`);
            if (typename === "int" || typename === "float") {
                buffer.writeLine(`readonly ${field.name}: (number${optional})[];`);
            } else if (typename === "string") {
                buffer.writeLine(`readonly ${field.name}: (string${optional})[];`);
            } else if (typename === "bool") {
                buffer.writeLine(`readonly ${field.name}: (boolean${optional})[];`);
            } else {
                buffer.writeLine(`readonly ${field.name}: (unknown${optional})[];`);
            }
        }
        buffer.unindent();
        buffer.writeLine(`}`);
        buffer.writeLine("");
    }
    return buffer.toString();
};

export const genLuaTypedef = (path: string, writer: string, maker?: ClassNameMaker) => {
    if (!maker) {
        maker = (className) => `xlsx.${writer}.${className}`;
    }
    const buffer = new StringBuffer(4);
    const workbook = get(path);
    const name = filename(path);
    for (const sheet of Object.values(workbook.sheets)) {
        const className = maker(toPascalCase(`${name}_${sheet.name}`));
        buffer.writeLine(`---file: ${path}`);
        buffer.writeLine(`---@class ${className}`);
        for (const field of sheet.fields) {
            if (!field.writers.includes(writer)) {
                continue;
            }
            const optional = field.typename.endsWith("?") ? "?" : "";
            let typename = field.typename.replaceAll("?", "");
            typename = convertors[typename].realtype ?? typename;
            const comment = field.comment.replaceAll(/[\r\n]+/g, " ");
            if (typename === "int") {
                buffer.writeLine(`---@field ${field.name}${optional} integer ${comment}`);
            } else if (typename === "float") {
                buffer.writeLine(`---@field ${field.name}${optional} number ${comment}`);
            } else if (typename === "string") {
                buffer.writeLine(`---@field ${field.name}${optional} string ${comment}`);
            } else if (typename === "bool") {
                buffer.writeLine(`---@field ${field.name}${optional} boolean ${comment}`);
            } else {
                buffer.writeLine(
                    `---@field ${field.name}${optional} ${maker(typename)} ${comment}`
                );
            }
        }
        buffer.writeLine("");
    }
    return buffer.toString();
};

export const genWorkbookTypedef = () => {
    const buffer = new StringBuffer(4);
    buffer.writeLine(`// AUTO GENERATED, DO NOT MODIFY!\n`);
    for (const path of Object.keys(files).sort()) {
        const workbook = get(path);
        const name = filename(path);
        for (const k of Object.keys(workbook.sheets).sort()) {
            const sheet = workbook.sheets[k];
            const className = toPascalCase(`${name}_${sheet.name}`);

            // row
            buffer.writeLine(`// file: ${path}`);
            buffer.writeLine(`export interface ${className} {`);
            buffer.indent();
            for (const field of sheet.fields) {
                const checker = field.checker.map((v) => v.def).join(";");
                const optional = field.typename.endsWith("?") ? "?" : "";
                const comment = field.comment.replaceAll(/[\r\n]+/g, " ");
                let typename = field.typename.replaceAll("?", "");
                typename = convertors[typename].realtype ?? typename;
                buffer.writeLine(`/**`);
                buffer.writeLine(
                    ` * ${comment} (location: ${field.refer}) (checker: ${checker || "x"}) ` +
                        `(writer: ${field.writers.join("|")})`
                );
                buffer.writeLine(` */`);
                if (typename === "int" || typename === "float") {
                    buffer.writeLine(`${field.name}: { v${optional}:number };`);
                } else if (typename === "string") {
                    buffer.writeLine(`${field.name}: { v${optional}:string };`);
                } else if (typename === "bool") {
                    buffer.writeLine(`${field.name}: { v${optional}:boolean };`);
                } else {
                    buffer.writeLine(`${field.name}: { v${optional}:unknown };`);
                }
            }
            buffer.unindent();
            buffer.writeLine(`}`);
            buffer.writeLine("");
        }
    }
    return buffer.toString();
};
