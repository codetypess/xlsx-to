import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import * as xlsx from "../index";

const makeCell = (value: xlsx.TValue, type: string, ref: string) => {
    return xlsx.makeCell(
        value,
        type,
        ref,
        value === null || value === undefined ? "" : String(value)
    );
};

const makeField = (index: number, name: string, typename: string = "string"): xlsx.Field => {
    return {
        index,
        name,
        typename,
        writers: [],
        checkers: [],
        comment: "",
        location: `A${index + 1}`,
        ignore: false,
    };
};

const makeSheet = (
    name: string,
    fields: Array<{ name: string; typename?: string }>,
    processors: { name: string; args: string[] }[] = []
) => {
    const data: xlsx.TObject = {};
    data["!type"] = xlsx.Type.Sheet;
    data["!name"] = name;
    return {
        name,
        ignore: false,
        processors,
        fields: fields.map((field, index) => makeField(index, field.name, field.typename)),
        data,
    } satisfies xlsx.Sheet;
};

const makeRow = (cells: Record<string, xlsx.TCell>) => {
    return {
        "!type": xlsx.Type.Row,
        ...cells,
    } as xlsx.TRow;
};

const runSheetTransformTests = () => {
    const ctx = new xlsx.Context("project-regression", "sheet-transforms");
    const workbook = new xlsx.Workbook(ctx, "test/regression/project-transform.xlsx");
    ctx.add(workbook);

    const define = makeSheet("define", [
        { name: "comment", typename: "string?" },
        { name: "enum", typename: "string?" },
        { name: "key", typename: "string?" },
        { name: "key1", typename: "string?" },
        { name: "key2", typename: "string?" },
        { name: "value", typename: "string?" },
        { name: "value_type", typename: "string" },
        { name: "value_comment", typename: "string?" },
    ]);

    define.data["1"] = makeRow({
        comment: makeCell("Currency", "string", "A2"),
        enum: makeCell("ProjectCurrency", "string", "B2"),
        key1: makeCell("Coin", "string", "D2"),
        key2: makeCell("Id", "string", "E2"),
        value: makeCell("1001", "string", "F2"),
        value_type: makeCell("int", "string", "G2"),
        value_comment: makeCell("Coin id", "string", "H2"),
    });
    define.data["2"] = makeRow({
        comment: makeCell("Status label", "string", "A3"),
        key: makeCell("Status", "string", "C3"),
        value: makeCell("enabled", "string", "F3"),
        value_type: makeCell("string", "string", "G3"),
    });

    workbook.add(define);

    const defined = xlsx.defineSheet(workbook, define);
    const coin = defined["Coin"] as xlsx.TObject;
    const coinId = coin["Id"] as xlsx.TCell;
    const status = defined["Status"] as xlsx.TCell;

    assert.equal(defined["!name"], "project-transform.define");
    assert.equal(defined["!type"], xlsx.Type.Define);
    assert.equal(coin["!enum"], "ProjectCurrency");
    assert.equal(coin["!comment"], "Currency");
    assert.equal(coinId.v, 1001);
    assert.equal(coinId["!comment"], "Coin id");
    assert.equal(status.v, "enabled");
    assert.equal(status["!comment"], "Status label");

    const duplicateDefine = makeSheet("duplicate-define", [
        { name: "key1", typename: "string" },
        { name: "value", typename: "string" },
        { name: "value_type", typename: "string" },
    ]);
    duplicateDefine.data["1"] = makeRow({
        key1: makeCell("Status", "string", "A2"),
        value: makeCell("enabled", "string", "B2"),
        value_type: makeCell("string", "string", "C2"),
    });
    duplicateDefine.data["2"] = makeRow({
        key1: makeCell("Status", "string", "A3"),
        value: makeCell("disabled", "string", "B3"),
        value_type: makeCell("string", "string", "C3"),
    });
    assert.throws(() => xlsx.defineSheet(workbook, duplicateDefine), /already defined/);

    const config = makeSheet("config", [
        { name: "cfg_key", typename: "string" },
        { name: "cfg_value", typename: "string" },
        { name: "cfg_type", typename: "string" },
        { name: "cfg_comment", typename: "string" },
    ]);
    config.data["1"] = makeRow({
        cfg_key: makeCell("maxRetry", "string", "A2"),
        cfg_value: makeCell("3", "string", "B2"),
        cfg_type: makeCell("int", "string", "C2"),
        cfg_comment: makeCell("Retry count", "string", "D2"),
    });

    const configured = xlsx.configSheet(
        workbook,
        config,
        "cfg_key",
        "cfg_value",
        "cfg_type",
        "cfg_comment"
    );
    const maxRetry = configured["maxRetry"] as xlsx.TCell;
    assert.equal(configured["!type"], xlsx.Type.Config);
    assert.equal(maxRetry.v, 3);
    assert.equal(maxRetry["!comment"], "Retry count");

    const duplicateConfig = makeSheet("duplicate-config", [
        { name: "key", typename: "string" },
        { name: "value", typename: "string" },
        { name: "value_type", typename: "string" },
        { name: "value_comment", typename: "string" },
    ]);
    duplicateConfig.data["1"] = makeRow({
        key: makeCell("lang", "string", "A2"),
        value: makeCell("zh", "string", "B2"),
        value_type: makeCell("string", "string", "C2"),
        value_comment: makeCell("Language", "string", "D2"),
    });
    duplicateConfig.data["2"] = makeRow({
        key: makeCell("lang", "string", "A3"),
        value: makeCell("en", "string", "B3"),
        value_type: makeCell("string", "string", "C3"),
        value_comment: makeCell("Language", "string", "D3"),
    });
    assert.throws(() => xlsx.configSheet(workbook, duplicateConfig), /already defined/);

    const resolved = xlsx.resolveDefineType<number>(workbook, workbook.path, define.name, "Coin");
    assert.deepEqual(resolved, { Id: 1001 });

    const lookup = makeSheet("lookup", [
        { name: "id", typename: "int" },
        { name: "group", typename: "string" },
        { name: "kind", typename: "string" },
        { name: "label", typename: "string" },
        { name: "value", typename: "int" },
    ]);
    lookup.data["1"] = makeRow({
        id: makeCell(1, "int", "A2"),
        group: makeCell("weapon", "string", "B2"),
        kind: makeCell("melee", "string", "C2"),
        label: makeCell("Sword", "string", "D2"),
        value: makeCell(100, "int", "E2"),
    });
    lookup.data["2"] = makeRow({
        id: makeCell(2, "int", "A3"),
        group: makeCell("weapon", "string", "B3"),
        kind: makeCell("range", "string", "C3"),
        label: makeCell("Bow", "string", "D3"),
        value: makeCell(200, "int", "E3"),
    });
    lookup.data["3"] = makeRow({
        id: makeCell(3, "int", "A4"),
        group: makeCell("armor", "string", "B4"),
        kind: makeCell("heavy", "string", "C4"),
        label: makeCell("Shield", "string", "D4"),
        value: makeCell(300, "int", "E4"),
    });
    workbook.add(lookup);

    const mappedObject = xlsx.mapSheet(workbook, lookup, "{label,value}", "group", "id");
    const weaponTwo = (mappedObject["weapon"] as xlsx.TObject)["2"] as xlsx.TObject;
    assert.equal((weaponTwo["label"] as xlsx.TCell).v as string, "Bow");
    assert.equal((weaponTwo["value"] as xlsx.TCell).v as number, 200);

    const mappedArray = xlsx.mapSheet(workbook, lookup, "[label,value]", "group", "id");
    const armorThree = ((mappedArray["armor"] as xlsx.TObject)["3"] as xlsx.TArray).map((cell) => {
        return (cell as xlsx.TCell).v;
    });
    assert.deepEqual(armorThree, ["Shield", 300]);

    const mappedRow = xlsx.mapSheet(workbook, lookup, "*", "group", "id");
    assert.equal(
        ((mappedRow["weapon"] as xlsx.TObject)["1"] as xlsx.TRow).label.v as string,
        "Sword"
    );
    assert.throws(() => xlsx.mapSheet(workbook, lookup, "invalid", "group"), /Invalid value query/);

    const collapsed = xlsx.collapseSheet(workbook, lookup, "group", "kind");
    assert.equal(((collapsed["weapon"] as xlsx.TObject)["melee"] as xlsx.TArray).length, 1);
    assert.equal(((collapsed["armor"] as xlsx.TObject)["heavy"] as xlsx.TArray).length, 1);

    const drops = makeSheet("drops", [
        { name: "id", typename: "int" },
        { name: "reward", typename: "int" },
        { name: "tag", typename: "string?" },
        { name: "note", typename: "string" },
    ]);
    drops.data["1"] = makeRow({
        id: makeCell(1, "int", "A2"),
        reward: makeCell(100, "int", "B2"),
        tag: makeCell("fire", "string", "C2"),
        note: makeCell("first", "string", "D2"),
    });
    drops.data["2"] = makeRow({
        id: makeCell(1, "int", "A3"),
        reward: makeCell(200, "int", "B3"),
        tag: makeCell(null, "string?", "C3"),
        note: makeCell("second", "string", "D3"),
    });
    drops.data["3"] = makeRow({
        id: makeCell(2, "int", "A4"),
        reward: makeCell(null, "int?", "B4"),
        tag: makeCell("ice", "string", "C4"),
        note: makeCell("third", "string", "D4"),
    });
    workbook.add(drops);

    const folded = xlsx.columnSheet(workbook, drops, "id", "reward", "tag");
    assert.deepEqual(
        ((folded["1"] as xlsx.TObject).reward as xlsx.TArray).map((cell) => (cell as xlsx.TCell).v),
        [100, 200]
    );
    assert.deepEqual(
        ((folded["1"] as xlsx.TObject).tag as xlsx.TArray).map((cell) => (cell as xlsx.TCell).v),
        ["fire"]
    );
    assert.equal((folded["1"] as xlsx.TObject).id, undefined);
    assert.equal(drops.fields[1]?.realtype, "int[]");
    assert.equal(drops.fields[2]?.realtype, "string[]");

    const simple = xlsx.simpleSheets(workbook, ["lookup"]);
    assert.equal(simple["lookup"], lookup.data);
    assert.equal(simple["drops"], undefined);

    const mergedWorkbook = new xlsx.Workbook(ctx, "test/regression/project-merge.xlsx");
    const first = makeSheet("first", [{ name: "id", typename: "int" }]);
    const second = makeSheet("second", [{ name: "id", typename: "int" }]);
    first.data["1"] = makeRow({ id: makeCell(1, "int", "A2") });
    second.data["2"] = makeRow({ id: makeCell(2, "int", "A2") });
    mergedWorkbook.add(first);
    mergedWorkbook.add(second);
    assert.deepEqual(Object.keys(xlsx.mergeSheets(mergedWorkbook)), ["1", "2"]);

    const duplicateMerge = new xlsx.Workbook(ctx, "test/regression/project-merge-dup.xlsx");
    const dupA = makeSheet("a", [{ name: "id", typename: "int" }]);
    const dupB = makeSheet("b", [{ name: "id", typename: "int" }]);
    dupA.data["1"] = makeRow({ id: makeCell(1, "int", "A2") });
    dupB.data["1"] = makeRow({ id: makeCell(2, "int", "A2") });
    duplicateMerge.add(dupA);
    duplicateMerge.add(dupB);
    assert.throws(() => xlsx.mergeSheets(duplicateMerge), /Duplicate key: 1/);
};

const runStringifyAndIndexerTests = () => {
    const sameCell = makeCell(9, "int", "A1");
    assert.equal(xlsx.convertValue(sameCell, "int"), sameCell);
    assert.deepEqual(xlsx.convertValue("[1, 2, 3]", "int[]"), [1, 2, 3]);
    assert.equal(xlsx.convertValue(makeCell("", "string?", "A2"), "int?").v, null);
    assert.equal(xlsx.convertValue("1", "bool"), true);
    assert.equal(xlsx.convertValue("x", "bool"), false);
    assert.equal(xlsx.convertValue("3.25", "float"), 3.25);
    assert.deepEqual(xlsx.convertValue('{foo: 1, nested: ["a", 2]}', "json"), {
        foo: 1,
        nested: ["a", 2],
    });
    assert.deepEqual(xlsx.convertValue('["a,b", "c"]', "string[]"), ["a,b", "c"]);
    assert.deepEqual(xlsx.convertValue("{1, 2, 3}", "table"), [1, 2, 3]);
    assert.deepEqual(xlsx.convertValue('{foo = 1, bar = {2, 3}, "tail"}', "table"), {
        1: "tail",
        foo: 1,
        bar: [2, 3],
    });
    assert.throws(() => xlsx.convertValue("12.5", "int"), /Convert value error: '12\.5'/);
    assert.throws(() => xlsx.convertValue("[1, 2]", "int[3]"), /type 'int\[3\]'/);
    assert.throws(() => xlsx.convertValue("{foo = bar-baz}", "table"), /type 'table'/);

    const status = {
        "!enum": "ProjectStatus",
        "!comment": "Status options",
        Active: makeCell(1, "int", "A3"),
        Disabled: makeCell(2, "int", "A4"),
    } satisfies xlsx.TObject;
    (status.Active as xlsx.TCell)["!comment"] = "Available";

    const stringifyData = {
        "!ignore": { hidden: true },
        42: makeCell("answer", "string", "B2"),
        title: makeCell("alpha\nbeta", "string", "B3"),
        meta: {
            "!comment": "Meta block",
            count: 1.25,
            enabled: true,
        } satisfies xlsx.TObject,
        status,
        hidden: makeCell("skip", "string", "B4"),
    } satisfies xlsx.TObject;

    const json = xlsx.stringifyJson(stringifyData, { indent: 2, precision: 2 });
    assert.match(json, /"42": "answer"/);
    assert.match(json, /"title": "alpha\\nbeta"/);
    assert.equal(json.includes("hidden"), false);

    const lua = xlsx.stringifyLua(stringifyData, { indent: 2, marshal: "return " });
    assert.match(lua, /return \{/);
    assert.match(lua, /\[42\] = "answer",/);
    assert.match(lua, /-- Meta block/);
    assert.match(lua, /---@enum ProjectStatus/);
    assert.equal(lua.includes("hidden"), false);

    const ts = xlsx.stringifyTs(stringifyData, { indent: 2, marshal: "export const catalog = " });
    assert.match(ts, /export type ProjectStatusKey/);
    assert.match(ts, /export enum ProjectStatus/);
    assert.match(ts, /export const catalog = \{/);
    assert.match(ts, /42: "answer",/);
    assert.match(ts, /title: "alpha\\nbeta",/);
    assert.equal(ts.includes("hidden:"), false);

    const tsType = xlsx.stringifyTsType(
        {
            info: {
                name: "Sword",
                count: 1,
            } satisfies xlsx.TObject,
            flags: [true, false] as xlsx.TArray,
            optional: undefined,
        } satisfies xlsx.TObject,
        { indent: 2, marshal: "export type Snapshot = " }
    );
    assert.match(tsType, /export type Snapshot = \{/);
    assert.match(tsType, /info: \{/);
    assert.match(tsType, /name: string;/);
    assert.match(tsType, /count: number;/);
    assert.match(tsType, /flags: \[/);
    assert.match(tsType, /boolean,/);
    assert.equal(tsType.includes("optional"), false);

    const ctx = new xlsx.Context("project-regression", "indexers");
    const workbook = new xlsx.Workbook(ctx, "test/regression/catalog.xlsx");
    ctx.add(workbook);

    const main = makeSheet("main", [
        { name: "id", typename: "int" },
        { name: "kind", typename: "string" },
        { name: "group", typename: "string" },
    ]);
    main.data["1"] = makeRow({
        id: makeCell(1, "int", "A2"),
        kind: makeCell("weapon", "string", "B2"),
        group: makeCell("equip", "string", "C2"),
    });
    main.data["2"] = makeRow({
        id: makeCell(2, "int", "A3"),
        kind: makeCell("weapon", "string", "B3"),
        group: makeCell("equip", "string", "C3"),
        $deprecated: makeCell(true, "bool", "D3"),
    });
    main.data["3"] = makeRow({
        id: makeCell(3, "int", "A4"),
        kind: makeCell("armor", "string", "B4"),
        group: makeCell("equip", "string", "C4"),
    });

    const extra = makeSheet("extra", [{ name: "id", typename: "int" }]);
    extra.data["10"] = makeRow({ id: makeCell(10, "int", "A2") });

    workbook.add(main);
    workbook.add(extra);

    const rowIndexer = new xlsx.RowIndexer(ctx, workbook.path, "main");
    assert.equal(rowIndexer.has("1"), true);
    assert.equal(rowIndexer.has("2"), false);
    assert.equal((rowIndexer.get("1") as xlsx.TRow).kind.v, "weapon");
    assert.equal((rowIndexer.get([{ key: "group", value: "equip" }]) as xlsx.TRow[]).length, 2);
    assert.equal((rowIndexer.get((row) => row.kind.v === "armor") as xlsx.TRow[]).length, 1);

    const columnIndexer = new xlsx.ColumnIndexer(ctx, workbook.path, "main", "id");
    assert.equal(columnIndexer.has(1), true);
    assert.equal(columnIndexer.has("1", [{ key: "kind", value: "weapon" }]), true);
    assert.equal(columnIndexer.get(3).length, 1);
    assert.equal(columnIndexer.get([{ key: "group", value: "equip" }]).length, 2);
    assert.equal(columnIndexer.get((row) => row.kind.v === "weapon").length, 1);

    const generated = xlsx.genWorkbookIndexer(ctx, (typename) => {
        return { type: typename, path: "./generated-types" };
    });
    assert.match(generated, /export class CatalogIndexer/);
    assert.match(generated, /static getRowIndexer\(ctx: Context, sheet: "main"/);
    assert.match(generated, /static getColumnIndexer\(ctx: Context, sheet: "extra"/);
    assert.match(
        generated,
        /return createRowIndexer\(ctx, "catalog\.xlsx", sheet, filter as Filter<unknown>\);/
    );
};

const runTsToZodTests = async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "typedsheet-zod-"));
    const defineDir = path.join(dir, "define");
    fs.mkdirSync(defineDir, { recursive: true });

    const inputPath = path.join(dir, "types.ts");
    const outputPath = path.join(dir, "types.schema.ts");
    const sharedPath = path.join(defineDir, "index.ts");
    const reexportPath = path.join(dir, "reexport.ts");

    fs.writeFileSync(
        sharedPath,
        ["export interface SharedType {", "    readonly value: number;", "}", ""].join("\n")
    );
    fs.writeFileSync(
        reexportPath,
        ["export interface ReExported {", "    readonly name: string;", "}", ""].join("\n")
    );
    fs.writeFileSync(
        inputPath,
        [
            'import type { SharedType } from "./define/index";',
            "",
            "interface LocalInline {",
            "    readonly note: string;",
            "}",
            "",
            'type LocalAlias = "a" | "b";',
            "",
            "export enum ProjectKind {",
            '    Alpha = "alpha",',
            '    Beta = "beta",',
            "}",
            "",
            "export interface Example {",
            "    readonly id: number;",
            "    readonly shared: SharedType;",
            "    readonly local: LocalInline;",
            "    readonly alias?: LocalAlias;",
            "    readonly kind: ProjectKind;",
            "}",
            "",
            "export type ExampleMap = Record<string, Example>;",
            "",
            'export * from "./reexport";',
            "",
        ].join("\n")
    );

    try {
        await xlsx.tsToZod(inputPath, outputPath);
        const output = fs.readFileSync(outputPath, "utf-8");
        assert.match(output, /import \{ z \} from "zod";/);
        assert.match(output, /import \{ ProjectKind \} from "\.\/types";/);
        assert.match(output, /import \{ sharedTypeSchema \} from "\.\/define\/index\.schema";/);
        assert.match(output, /export const projectKindSchema = z\.enum\(ProjectKind\);/);
        assert.match(output, /export const exampleSchema = z\.object\(\{/);
        assert.match(output, /shared: sharedTypeSchema/);
        assert.match(output, /local: z\.object\(\{/);
        assert.match(output, /note: z\.string\(\)/);
        assert.match(
            output,
            /alias: z\.union\(\[z\.literal\("a"\), z\.literal\("b"\)\]\)\.optional\(\)/
        );
        assert.match(output, /kind: projectKindSchema/);
        assert.match(
            output,
            /export const exampleMapSchema = z\.record\(z\.string\(\), exampleSchema\);/
        );
        assert.match(output, /export \* from "\.\/reexport\.schema";/);
    } finally {
        fs.rmSync(dir, { recursive: true, force: true });
    }
};

export const runProjectRegressionTests = async () => {
    runSheetTransformTests();
    runStringifyAndIndexerTests();
    await runTsToZodTests();
};
