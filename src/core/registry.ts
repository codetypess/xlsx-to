import type { CheckerParser, Converter, Processor, Writer } from "./contracts.js";
import { assert } from "./errors.js";

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

export const settings = {
    suppressCheckers: new Set<string>(),
    suppressProcessors: new Set<string>(),
    suppressWriters: new Set<string>(),
} as const;

export const DEFAULT_WRITER = "__xlsx_default_writer__";
export const DEFAULT_TAG = "__xlsx_default_tag__";
export const checkerParsers: Record<string, CheckerParser> = {};
export const converters: Record<string, Converter> = {};
export const processors: Record<string, ProcessorType> = {};
export const writers: Record<string, Writer> = {};

export function registerType(typename: string, converter: Converter): void {
    assert(typeof converter === "function", `Converter must be a function: '${typename}'`);
    if (converters[typename]) {
        console.warn(`Overwrite previous registered converter '${typename}'`);
    }
    converters[typename] = converter;
}

export const registerChecker = (name: string, parser: CheckerParser) => {
    if (checkerParsers[name]) {
        console.warn(`Overwrite previous registered checker parser '${name}'`);
    }
    checkerParsers[name] = parser;
};

export const suppressChecker = (name: string) => {
    settings.suppressCheckers.add(name);
};

export const suppressAllCheckers = () => {
    Object.keys(checkerParsers).forEach(suppressChecker);
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

export const suppressProcessor = (name: string) => {
    settings.suppressProcessors.add(name);
};

export const registerWriter = (name: string, writer: Writer) => {
    if (writers[name]) {
        console.warn(`Overwrite previous registered writer '${name}'`);
    }
    writers[name] = writer;
};

export const suppressWriter = (name: string) => {
    settings.suppressWriters.add(name);
};
