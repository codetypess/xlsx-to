import { assert } from "./errors.js";
import type { CheckerParser, Convertor, Processor, Writer } from "../xlsx.js";

export type ProcessorStage =
    | "after-read"
    | "pre-parse"
    | "after-parse"
    | "pre-check"
    | "after-check"
    | "pre-stringify"
    | "stringify"
    | "after-stringify";

export type ProcessorOption = {
    /** Automatically added to every workbook. */
    readonly required: boolean;
    /** The priority of the processor, higher value means lower priority */
    readonly priority: number;
    readonly stage: ProcessorStage;
};

export type ProcessorType = {
    readonly name: string;
    readonly option: ProcessorOption;
    readonly exec: Processor;
};

export const options = {
    suppressCheckers: [] as string[],
    suppressProcessors: [] as string[],
    suppressWriters: [] as string[],
};

export const DEFAULT_WRITER = "__xlsx_default_writer__";
export const DEFAULT_TAG = "__xlsx_default_tag__";
export const checkerParsers: Record<string, CheckerParser> = {};
export const convertors: Record<string, Convertor> = {};
export const processors: Record<string, ProcessorType> = {};
export const writers: Record<string, Writer> = {};

export function registerType(typename: string, convertor: Convertor): void {
    assert(typeof convertor === "function", `Convertor must be a function: '${typename}'`);
    if (convertors[typename]) {
        console.warn(`Overwrite previous registered convertor '${typename}'`);
    }
    convertors[typename] = convertor;
}

export const registerChecker = (name: string, parser: CheckerParser) => {
    if (checkerParsers[name]) {
        console.warn(`Overwrite previous registered checker parser '${name}'`);
    }
    checkerParsers[name] = parser;
};

export const registerProcessor = (
    name: string,
    processor: Processor,
    option?: Partial<ProcessorOption>
) => {
    if (processors[name]) {
        console.warn(`Overwrite previous registered processor '${name}'`);
    }
    processors[name] = {
        name,
        option: {
            required: option?.required ?? false,
            stage: option?.stage ?? "stringify",
            priority: option?.priority ?? 0,
        },
        exec: processor,
    };
};

export const registerWriter = (name: string, writer: Writer) => {
    if (writers[name]) {
        console.warn(`Overwrite previous registered writer '${name}'`);
    }
    writers[name] = writer;
};
