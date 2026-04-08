import {
    ExprCheckerParser,
    FollowCheckerParser,
    IndexCheckerParser,
    OneOfCheckerParser,
    RangeCheckerParser,
    ReferCheckerParser,
    SheetCheckerParser,
    SizeCheckerParser,
    UniqueCheckerParser,
} from "../builtins/checkers.js";
import {
    boolConverter,
    floatConverter,
    intConverter,
    jsonConverter,
    stringConverter,
} from "../builtins/converters.js";
import {
    AutoRegisterProcessor,
    CollapseProcessor,
    ColumnProcessor,
    ConfigProcessor,
    DefineProcessor,
    GenTypeProcessor,
    MapProcessor,
    mergeSheets,
    registerStringifyRule,
    simpleSheets,
    StringifyProcessor,
    TypedefProcessor,
    TypedefWriteProcessor,
} from "../builtins/processors.js";
import { BuiltinChecker } from "../core/contracts.js";
import { registerChecker, registerProcessor, registerType } from "../core/registry.js";
import { tableConverter } from "../core/table.js";

let registered = false;

export const registerBuiltins = () => {
    if (registered) {
        return;
    }
    registered = true;

    registerType("auto", intConverter);
    registerType("bool", boolConverter);
    registerType("float", floatConverter);
    registerType("int", intConverter);
    registerType("json", jsonConverter);
    registerType("string", stringConverter);
    registerType("table", tableConverter);

    registerChecker(BuiltinChecker.Expr, ExprCheckerParser);
    registerChecker(BuiltinChecker.Follow, FollowCheckerParser);
    registerChecker(BuiltinChecker.Index, IndexCheckerParser);
    registerChecker(BuiltinChecker.OneOf, OneOfCheckerParser);
    registerChecker(BuiltinChecker.Range, RangeCheckerParser);
    registerChecker(BuiltinChecker.Refer, ReferCheckerParser);
    registerChecker(BuiltinChecker.Sheet, SheetCheckerParser);
    registerChecker(BuiltinChecker.Size, SizeCheckerParser);
    registerChecker(BuiltinChecker.Unique, UniqueCheckerParser);

    registerProcessor("define", DefineProcessor, { stage: "pre-stringify" });
    registerProcessor("config", ConfigProcessor, { stage: "pre-stringify", priority: 800 });
    registerProcessor("map", MapProcessor, { stage: "pre-stringify", priority: 800 });
    registerProcessor("collapse", CollapseProcessor, { stage: "pre-stringify", priority: 800 });
    registerProcessor("column", ColumnProcessor, { stage: "pre-stringify", priority: 800 });
    registerProcessor("typedef", TypedefProcessor, { stage: "after-read" });
    registerProcessor("typedef-write", TypedefWriteProcessor, {
        stage: "pre-stringify",
        priority: -100,
    });
    registerProcessor("stringify", StringifyProcessor, {
        stage: "stringify",
        priority: 900,
        required: true,
    });
    registerProcessor("gen-type", GenTypeProcessor, {
        stage: "stringify",
        priority: 999,
        required: true,
    });
    registerProcessor("auto-register", AutoRegisterProcessor, {
        required: true,
        stage: "after-read",
        priority: 999,
    });

    registerStringifyRule("merge", mergeSheets);
    registerStringifyRule("simple", simpleSheets);
};
