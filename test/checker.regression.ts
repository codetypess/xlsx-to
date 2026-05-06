import assert from "node:assert/strict";
import * as xlsx from "../index";
import {
    ExprCheckerParser,
    FollowCheckerParser,
    IndexCheckerParser,
    RangeCheckerParser,
    ReferCheckerParser,
    SheetCheckerParser,
    SizeCheckerParser,
    UniqueCheckerParser,
} from "../src/builtins/checkers";
import type { CheckerType } from "../src/core/contracts";
import { loadBody, parseChecker } from "../src/core/parser";
import { performChecker, resolveChecker } from "../src/core/pipeline";

const makeCell = (value: xlsx.TValue, type: string, ref: string) => {
    let source = "";
    if (value !== null && value !== undefined) {
        source = typeof value === "object" ? JSON.stringify(value) : String(value);
    }
    return xlsx.makeCell(value, type, ref, source);
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

const makeSheet = (name: string, fields: Array<{ name: string; typename?: string }>) => {
    const data: xlsx.TObject = {};
    data["!type"] = xlsx.Type.Sheet;
    data["!name"] = name;
    return {
        name,
        ignore: false,
        processors: [],
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

const clearAllContexts = () => {
    for (const ctx of xlsx.getContexts().slice()) {
        xlsx.removeContext(ctx);
    }
};

const runParseCheckerTests = () => {
    assert.deepEqual(parseChecker("source.xlsx", "main", "A1", 0, "x"), []);
    assert.deepEqual(parseChecker("source.xlsx", "main", "A1", 0, "!!skip"), []);

    const forceUnique = parseChecker("source.xlsx", "main", "A1", 1, "!@unique")[0]!;
    assert.equal(forceUnique.name, xlsx.BuiltinChecker.Unique);
    assert.equal(forceUnique.force, true);
    assert.deepEqual(forceUnique.args, [""]);

    const range = parseChecker("source.xlsx", "main", "A1", 1, "[1,2,3]")[0]!;
    assert.equal(range.name, xlsx.BuiltinChecker.Range);
    assert.deepEqual(range.args, ["[1,2,3]"]);

    const sheet = parseChecker("source.xlsx", "main", "A1", 1, "target#")[0]!;
    assert.equal(sheet.name, xlsx.BuiltinChecker.Sheet);
    assert.deepEqual(sheet.args, ["source.xlsx", "main", "", "", "target.xlsx"]);

    const index = parseChecker(
        "source.xlsx",
        "main",
        "A1",
        1,
        "$.items[*].id&kind=A==target#refs.id&group=B"
    )[0]!;
    assert.equal(index.name, xlsx.BuiltinChecker.Index);
    assert.deepEqual(index.args, [
        "source.xlsx",
        "main",
        ".items[*].id",
        "kind=A",
        "target.xlsx",
        "refs",
        "id",
        "group=B",
    ]);
};

const runReferBindingTests = () => {
    clearAllContexts();
    try {
        const ctx = xlsx.addContext(new xlsx.Context(xlsx.DEFAULT_WRITER, xlsx.DEFAULT_TAG));
        const workbook = new xlsx.Workbook(ctx, "refer-binding.xlsx");
        const sheet = makeSheet("main", [
            { name: "id", typename: "int" },
            { name: "values", typename: "int[]?" },
            { name: "rules", typename: "string?" },
        ]);
        (sheet.fields[1]!.checkers as CheckerType[]).push(
            ...parseChecker(workbook.path, sheet.name, "B4", 1, "@refer(rules)")
        );
        workbook.add(sheet);
        ctx.add(workbook);

        const cells = new Map<string, string | number>([
            ["A1", ""],
            ["A6", "1"],
            ["B6", "[10,20]"],
            ["C6", "@size(2)"],
        ]);
        const rawSheet = {
            name: "main",
            rowCount: 6,
            columnCount: 3,
            getCell(row: number, col: number) {
                const column = String.fromCharCode(64 + col);
                return cells.get(`${column}${row}`) ?? "";
            },
        };
        const rawWorkbook = {
            getSheets() {
                return [rawSheet];
            },
        };

        loadBody(workbook.path, rawWorkbook as never);

        const referChecker = sheet.fields[1]!.checkers[0] as CheckerType;
        assert.equal(Array.isArray(referChecker.refers["B6"]), true);
        assert.equal(referChecker.refers["A6"], undefined);
        assert.equal(referChecker.refers["B6"]?.[0]?.source, "@size(2)");
    } finally {
        clearAllContexts();
    }
};

const runSingleCheckerParserTests = () => {
    assert.throws(() => SizeCheckerParser({} as xlsx.Context, "oops"), /Invalid length/);

    {
        const checker = SizeCheckerParser({} as xlsx.Context, "2");
        assert.equal(
            checker({
                workbook: {} as xlsx.Workbook,
                sheet: {} as xlsx.Sheet,
                field: makeField(0, "items", "int[]"),
                row: makeRow({}),
                cell: makeCell([1, 2], "int[]", "A2"),
                errors: [],
            }),
            true
        );
        assert.equal(
            checker({
                workbook: {} as xlsx.Workbook,
                sheet: {} as xlsx.Sheet,
                field: makeField(0, "items", "int[]"),
                row: makeRow({}),
                cell: makeCell(1, "int", "A3"),
                errors: [],
            }),
            false
        );
    }

    {
        const checker = ExprCheckerParser({} as xlsx.Context, "missing.value > 0");
        const errors: string[] = [];
        const ok = checker({
            workbook: {} as xlsx.Workbook,
            sheet: {} as xlsx.Sheet,
            field: makeField(0, "value", "int"),
            row: makeRow({ value: makeCell(1, "int", "A2") }),
            cell: makeCell(1, "int", "A2"),
            errors,
        });
        assert.equal(ok, false);
        assert.equal(errors.length, 1);
        assert.match(errors[0]!, /^Expression error:/);
    }

    {
        const checker = FollowCheckerParser({} as xlsx.Context, "required");
        assert.equal(
            checker({
                workbook: {} as xlsx.Workbook,
                sheet: {} as xlsx.Sheet,
                field: makeField(0, "value", "int?"),
                row: makeRow({
                    required: makeCell(1, "int", "B2"),
                }),
                cell: makeCell(2, "int", "A2"),
                errors: [],
            }),
            true
        );
        assert.equal(
            checker({
                workbook: {} as xlsx.Workbook,
                sheet: {} as xlsx.Sheet,
                field: makeField(0, "value", "int?"),
                row: makeRow({
                    required: makeCell(1, "int", "B3"),
                }),
                cell: makeCell(null, "int?", "A3"),
                errors: [],
            }),
            false
        );
        assert.equal(
            checker({
                workbook: {} as xlsx.Workbook,
                sheet: {} as xlsx.Sheet,
                field: makeField(0, "value", "int?"),
                row: makeRow({
                    required: makeCell(null, "int?", "B4"),
                }),
                cell: makeCell(null, "int?", "A4"),
                errors: [],
            }),
            true
        );
    }

    assert.throws(() => RangeCheckerParser({} as xlsx.Context, "[1,"), /Invalid range/);

    {
        const checker = RangeCheckerParser({} as xlsx.Context, "[1,2,3]");
        assert.equal(
            checker({
                workbook: {} as xlsx.Workbook,
                sheet: {} as xlsx.Sheet,
                field: makeField(0, "id", "int"),
                row: makeRow({}),
                cell: makeCell(2, "int", "A2"),
                errors: [],
            }),
            true
        );
        assert.equal(
            checker({
                workbook: {} as xlsx.Workbook,
                sheet: {} as xlsx.Sheet,
                field: makeField(0, "id", "int"),
                row: makeRow({}),
                cell: makeCell(4, "int", "A3"),
                errors: [],
            }),
            false
        );
    }

    {
        const checker = ReferCheckerParser({} as xlsx.Context, "anything");
        const errors: string[] = [];
        const ok = checker({
            workbook: {} as xlsx.Workbook,
            sheet: {} as xlsx.Sheet,
            field: makeField(0, "id", "int"),
            row: makeRow({}),
            cell: makeCell(1, "int", "A2"),
            errors,
        });
        assert.equal(ok, true);
        assert.deepEqual(errors, []);
    }
};

const runIndexAndSheetCheckerTests = () => {
    const ctx = new xlsx.Context("checker-regression", "single-scenarios");
    const sourceWorkbook = new xlsx.Workbook(ctx, "source.xlsx");
    const targetWorkbook = new xlsx.Workbook(ctx, "target.xlsx");
    const sourceSheet = makeSheet("main", [
        { name: "payload", typename: "json" },
        { name: "kind", typename: "string" },
    ]);
    const itemsSheet = makeSheet("items", [
        { name: "id", typename: "int" },
        { name: "code", typename: "string" },
        { name: "group", typename: "string" },
    ]);
    const refsSheet = makeSheet("refs", [
        { name: "id", typename: "int" },
        { name: "group", typename: "string" },
    ]);
    const alphaSheet = makeSheet("alpha", [{ name: "id", typename: "int" }]);
    const betaSheet = makeSheet("beta", [{ name: "id", typename: "int" }]);

    itemsSheet.data["1"] = makeRow({
        id: makeCell(1, "int", "A2"),
        code: makeCell("sword", "string", "B2"),
        group: makeCell("A", "string", "C2"),
    });
    itemsSheet.data["2"] = makeRow({
        id: makeCell(2, "int", "A3"),
        code: makeCell("shield", "string", "B3"),
        group: makeCell("B", "string", "C3"),
    });
    refsSheet.data["1"] = makeRow({
        id: makeCell(1, "int", "A2"),
        group: makeCell("A", "string", "B2"),
    });

    ctx.add(sourceWorkbook);
    ctx.add(targetWorkbook);
    sourceWorkbook.add(sourceSheet);
    targetWorkbook.add(itemsSheet);
    targetWorkbook.add(refsSheet);
    targetWorkbook.add(alphaSheet);
    targetWorkbook.add(betaSheet);

    {
        const checker = IndexCheckerParser(
            ctx,
            "source.xlsx",
            "main",
            "[.]",
            "",
            "target.xlsx",
            "items",
            "code",
            ""
        );
        const errors: string[] = [];
        const ok = checker({
            workbook: sourceWorkbook,
            sheet: sourceSheet,
            field: sourceSheet.fields[0]!,
            row: makeRow({
                payload: makeCell({ sword: 1, shield: 2 }, "json", "A2"),
                kind: makeCell("A", "string", "B2"),
            }),
            cell: makeCell({ sword: 1, shield: 2 }, "json", "A2"),
            errors,
        });
        assert.equal(ok, true);
        assert.deepEqual(errors, []);
    }

    {
        const checker = IndexCheckerParser(
            ctx,
            "source.xlsx",
            "main",
            "[*].id",
            "",
            "target.xlsx",
            "items",
            "id",
            ""
        );
        const errors: string[] = [];
        const ok = checker({
            workbook: sourceWorkbook,
            sheet: sourceSheet,
            field: sourceSheet.fields[0]!,
            row: makeRow({
                payload: makeCell([{ id: 1 }, { id: 2 }], "json", "A3"),
                kind: makeCell("A", "string", "B3"),
            }),
            cell: makeCell([{ id: 1 }, { id: 2 }], "json", "A3"),
            errors,
        });
        assert.equal(ok, true);
        assert.deepEqual(errors, []);
    }

    {
        const checker = IndexCheckerParser(
            ctx,
            "source.xlsx",
            "main",
            "",
            "kind=B",
            "target.xlsx",
            "items",
            "id",
            ""
        );
        const errors: string[] = [];
        const ok = checker({
            workbook: sourceWorkbook,
            sheet: sourceSheet,
            field: sourceSheet.fields[0]!,
            row: makeRow({
                payload: makeCell(999, "int", "A4"),
                kind: makeCell("A", "string", "B4"),
            }),
            cell: makeCell(999, "int", "A4"),
            errors,
        });
        assert.equal(ok, true);
        assert.deepEqual(errors, []);
    }

    {
        const checker = IndexCheckerParser(
            ctx,
            "source.xlsx",
            "main",
            "[1]?",
            "",
            "target.xlsx",
            "items",
            "id",
            ""
        );
        const errors: string[] = [];
        const ok = checker({
            workbook: sourceWorkbook,
            sheet: sourceSheet,
            field: sourceSheet.fields[0]!,
            row: makeRow({
                payload: makeCell([1], "int[]", "A5"),
                kind: makeCell("A", "string", "B5"),
            }),
            cell: makeCell([1], "int[]", "A5"),
            errors,
        });
        assert.equal(ok, true);
        assert.deepEqual(errors, []);
    }

    {
        const checker = IndexCheckerParser(
            ctx,
            "source.xlsx",
            "main",
            ".meta",
            "",
            "target.xlsx",
            "items",
            "id",
            ""
        );
        const errors: string[] = [];
        const ok = checker({
            workbook: sourceWorkbook,
            sheet: sourceSheet,
            field: sourceSheet.fields[0]!,
            row: makeRow({
                payload: makeCell({ meta: { bad: true } }, "json", "A6"),
                kind: makeCell("A", "string", "B6"),
            }),
            cell: makeCell({ meta: { bad: true } }, "json", "A6"),
            errors,
        });
        assert.equal(ok, false);
        assert.ok(errors.some((entry) => entry.startsWith("data type error:")));
    }

    {
        const checker = SheetCheckerParser(ctx, "source.xlsx", "main", "[.]", "", "target.xlsx");
        const errors: string[] = [];
        const ok = checker({
            workbook: sourceWorkbook,
            sheet: sourceSheet,
            field: sourceSheet.fields[0]!,
            row: makeRow({
                payload: makeCell({ alpha: 1, beta: 2 }, "json", "A7"),
                kind: makeCell("A", "string", "B7"),
            }),
            cell: makeCell({ alpha: 1, beta: 2 }, "json", "A7"),
            errors,
        });
        assert.equal(ok, true);
        assert.deepEqual(errors, []);
    }
};

const runUniqueCheckerTests = () => {
    const ctx = new xlsx.Context("checker-regression", "unique-scenarios");
    const workbook = new xlsx.Workbook(ctx, "unique.xlsx");
    const sheet = makeSheet("main", [
        { name: "id", typename: "int" },
        { name: "name", typename: "string" },
    ]);

    sheet.data["1"] = makeRow({
        id: makeCell(1, "int", "A2"),
        name: makeCell("Sword", "string", "B2"),
    });
    sheet.data["2"] = makeRow({
        id: makeCell(1, "int", "A3"),
        name: makeCell("Sword Copy", "string", "B3"),
    });
    sheet.data["3"] = makeRow({
        id: makeCell(2, "int", "A4"),
        name: makeCell("Shield", "string", "B4"),
    });

    ctx.add(workbook);
    workbook.add(sheet);

    const checker = UniqueCheckerParser(ctx, "");

    {
        const errors: string[] = [];
        const ok = checker({
            workbook,
            sheet,
            field: sheet.fields[0]!,
            row: sheet.data["1"] as xlsx.TRow,
            cell: (sheet.data["1"] as xlsx.TRow).id,
            errors,
        });
        assert.equal(ok, false);
        assert.ok(errors.includes("unique error: location=A3"));
    }

    {
        const errors: string[] = [];
        const ok = checker({
            workbook,
            sheet,
            field: sheet.fields[0]!,
            row: sheet.data["3"] as xlsx.TRow,
            cell: (sheet.data["3"] as xlsx.TRow).id,
            errors,
        });
        assert.equal(ok, true);
        assert.deepEqual(errors, []);
    }

    {
        const objectValueSheet = makeSheet("object_value", [{ name: "id", typename: "json" }]);
        objectValueSheet.data["1"] = makeRow({
            id: makeCell({ nested: true }, "json", "A2"),
        });
        workbook.add(objectValueSheet);
        const errors: string[] = [];
        const ok = checker({
            workbook,
            sheet: objectValueSheet,
            field: objectValueSheet.fields[0]!,
            row: objectValueSheet.data["1"] as xlsx.TRow,
            cell: (objectValueSheet.data["1"] as xlsx.TRow).id,
            errors,
        });
        assert.equal(ok, false);
        assert.deepEqual(errors, ["data type error: type=object"]);
    }
};

const runForceCheckerPipelineTests = () => {
    clearAllContexts();
    try {
        {
            const ctx = xlsx.addContext(new xlsx.Context("client", "checker-force-skip"));
            const workbook = new xlsx.Workbook(ctx, "force-skip.xlsx");
            const sheet = makeSheet("main", [
                { name: "id", typename: "int" },
                { name: "values", typename: "int[]?" },
            ]);
            (sheet.fields[1]!.checkers as CheckerType[]).push(
                ...parseChecker(workbook.path, sheet.name, "B1", 1, "@size(1)")
            );
            sheet.data["1"] = makeRow({
                id: makeCell(1, "int", "A2"),
                values: makeCell(null, "int[]?", "B2"),
            });
            ctx.add(workbook);
            workbook.add(sheet);
            resolveChecker();
            assert.doesNotThrow(() => performChecker());
        }

        clearAllContexts();

        {
            const ctx = xlsx.addContext(new xlsx.Context("client", "checker-force-enforced"));
            const workbook = new xlsx.Workbook(ctx, "force-enforced.xlsx");
            const sheet = makeSheet("main", [
                { name: "id", typename: "int" },
                { name: "values", typename: "int[]?" },
            ]);
            (sheet.fields[1]!.checkers as CheckerType[]).push(
                ...parseChecker(workbook.path, sheet.name, "B1", 1, "!@size(1)")
            );
            sheet.data["1"] = makeRow({
                id: makeCell(1, "int", "A2"),
                values: makeCell(null, "int[]?", "B2"),
            });
            ctx.add(workbook);
            workbook.add(sheet);
            resolveChecker();
            assert.throws(() => performChecker(), /checker: @size\(1\)/);
        }
    } finally {
        clearAllContexts();
    }
};

const runReferCheckerPipelineTests = () => {
    clearAllContexts();
    try {
        {
            const ctx = xlsx.addContext(new xlsx.Context("client", "checker-refer-success"));
            const workbook = new xlsx.Workbook(ctx, "refer-success.xlsx");
            const sheet = makeSheet("main", [
                { name: "id", typename: "int" },
                { name: "values", typename: "int[]?" },
                { name: "rules", typename: "string?" },
            ]);
            (sheet.fields[1]!.checkers as CheckerType[]).push(
                ...parseChecker(workbook.path, sheet.name, "B1", 1, "@refer(rules)")
            );
            const referChecker = sheet.fields[1]!.checkers[0] as CheckerType;
            referChecker.refers["B2"] = parseChecker(
                workbook.path,
                sheet.name,
                "C2",
                2,
                "@size(2)"
            );
            sheet.data["1"] = makeRow({
                id: makeCell(1, "int", "A2"),
                values: makeCell([10, 20], "int[]", "B2"),
                rules: makeCell("@size(2)", "string", "C2"),
            });
            ctx.add(workbook);
            workbook.add(sheet);
            resolveChecker();
            assert.doesNotThrow(() => performChecker());
        }

        clearAllContexts();

        {
            const ctx = xlsx.addContext(new xlsx.Context("client", "checker-refer-empty-rule"));
            const workbook = new xlsx.Workbook(ctx, "refer-empty-rule.xlsx");
            const sheet = makeSheet("main", [
                { name: "id", typename: "int" },
                { name: "values", typename: "int[]?" },
                { name: "rules", typename: "string?" },
            ]);
            (sheet.fields[1]!.checkers as CheckerType[]).push(
                ...parseChecker(workbook.path, sheet.name, "B1", 1, "@refer(rules)")
            );
            sheet.data["1"] = makeRow({
                id: makeCell(1, "int", "A2"),
                values: makeCell([10], "int[]", "B2"),
                rules: makeCell("", "string?", "C2"),
            });
            ctx.add(workbook);
            workbook.add(sheet);
            resolveChecker();
            assert.doesNotThrow(() => performChecker());
        }

        clearAllContexts();

        {
            const ctx = xlsx.addContext(new xlsx.Context("client", "checker-refer-force-null"));
            const workbook = new xlsx.Workbook(ctx, "refer-force-null.xlsx");
            const sheet = makeSheet("main", [
                { name: "id", typename: "int" },
                { name: "values", typename: "int[]?" },
                { name: "rules", typename: "string?" },
            ]);
            (sheet.fields[1]!.checkers as CheckerType[]).push(
                ...parseChecker(workbook.path, sheet.name, "B1", 1, "@refer(rules)")
            );
            const referChecker = sheet.fields[1]!.checkers[0] as CheckerType;
            referChecker.refers["B2"] = parseChecker(
                workbook.path,
                sheet.name,
                "C2",
                2,
                "!@size(1)"
            );
            sheet.data["1"] = makeRow({
                id: makeCell(1, "int", "A2"),
                values: makeCell(null, "int[]?", "B2"),
                rules: makeCell("!@size(1)", "string", "C2"),
            });
            ctx.add(workbook);
            workbook.add(sheet);
            resolveChecker();
            assert.throws(() => performChecker(), /checker: @size\(1\)/);
        }
    } finally {
        clearAllContexts();
    }
};

export const runCheckerRegressionTests = async () => {
    runParseCheckerTests();
    runReferBindingTests();
    runSingleCheckerParserTests();
    runIndexAndSheetCheckerTests();
    runUniqueCheckerTests();
    runForceCheckerPipelineTests();
    runReferCheckerPipelineTests();
};
