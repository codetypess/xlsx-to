import assert from "node:assert/strict";
import * as xlsx from "../index";
import { error as raiseError, assert as runtimeAssert, trace } from "../src/core/errors";
import { loadBody, loadHeader } from "../src/core/parser";
import { copyWorkbook } from "../src/core/pipeline";

const makeCell = (value: xlsx.TValue, type?: string, ref: string = "A1") => {
    const source = value === null || value === undefined ? "" : String(value);
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

const makeSheet = (
    name: string,
    fields: Array<{ name: string; typename?: string; writers?: string[] }>
) => {
    const data: xlsx.TObject = {};
    data["!type"] = xlsx.Type.Sheet;
    data["!name"] = name;
    return {
        name,
        ignore: false,
        processors: [],
        fields: fields.map((field, index) => ({
            ...makeField(index, field.name, field.typename),
            writers: field.writers ?? [],
        })),
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

const withPatchedConsole = <T>(method: "warn" | "error", run: (messages: string[]) => T): T => {
    const original = console[method];
    const messages: string[] = [];
    console[method] = ((...args: unknown[]) => {
        messages.push(args.map((arg) => String(arg)).join(" "));
    }) as (typeof console)[typeof method];
    try {
        return run(messages);
    } finally {
        console[method] = original;
    }
};

const withSuppressionSnapshot = (run: () => void) => {
    const checkerSnapshot = new Set(xlsx.settings.suppressCheckers);
    const processorSnapshot = new Set(xlsx.settings.suppressProcessors);
    const writerSnapshot = new Set(xlsx.settings.suppressWriters);
    try {
        run();
    } finally {
        xlsx.settings.suppressCheckers.clear();
        xlsx.settings.suppressProcessors.clear();
        xlsx.settings.suppressWriters.clear();
        checkerSnapshot.forEach((name) => xlsx.settings.suppressCheckers.add(name));
        processorSnapshot.forEach((name) => xlsx.settings.suppressProcessors.add(name));
        writerSnapshot.forEach((name) => xlsx.settings.suppressWriters.add(name));
    }
};

const makeRawSheet = (
    name: string,
    columnCount: number,
    rowCount: number,
    cells: Record<string, string | number | null>
) => {
    const getAddress = (row: number, col: number) => {
        const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        let index = col;
        let label = "";
        while (index > 0) {
            const remainder = (index - 1) % 26;
            label = letters[remainder] + label;
            index = Math.floor((index - 1) / 26);
        }
        return `${label}${row}`;
    };

    return {
        name,
        columnCount,
        rowCount,
        getCell(row: number, col: number) {
            return cells[getAddress(row, col)] ?? "";
        },
    };
};

const runErrorsAndValueTests = () => {
    {
        let failure: Error | undefined;
        try {
            using _ = trace("outer");
            raiseError("boom");
        } catch (error) {
            failure = error as Error;
        }
        assert(failure);
        assert.match(failure.message, /^boom/);
        assert.match(failure.message, /-> outer/);
    }

    {
        let failure: Error | undefined;
        try {
            using _ = trace("guard");
            runtimeAssert(false, "failed condition");
        } catch (error) {
            failure = error as Error;
        }
        assert(failure);
        assert.match(failure.message, /failed condition/);
        assert.match(failure.message, /-> guard/);
    }

    {
        const src = {
            "!type": xlsx.Type.Define,
            "!name": "tagged",
            plain: 1,
        } satisfies xlsx.TObject;
        const dest: xlsx.TObject = { plain: 2 };
        xlsx.copyTag(src, dest);
        assert.equal(dest["!type"], xlsx.Type.Define);
        assert.equal(dest["!name"], "tagged");
        assert.equal(dest.plain, 2);
    }

    assert.equal(xlsx.typeOf(1), "number");
    assert.equal(xlsx.typeOf(makeCell(1, "int", "A2")), xlsx.Type.Cell);
    assert.equal(xlsx.checkType<number>(1, "number"), 1);

    withPatchedConsole("error", (messages) => {
        assert.throws(() => xlsx.checkType(1, xlsx.Type.Cell), /Expect type 'xlsx\.type\.cell'/);
        assert.equal(messages.length, 1);
    });

    assert.equal(xlsx.isNull(null), true);
    assert.equal(xlsx.isNull(makeCell(null, "int?", "A3")), true);
    assert.equal(xlsx.isNotNull(makeCell(1, "int", "A4")), true);

    const tagged = {} as xlsx.TObject;
    xlsx.ignoreField(tagged, "hidden", true);
    assert.deepEqual(tagged["!ignore"], { hidden: true });

    assert.equal(xlsx.toString(makeCell("  trim  ", "string", "A5")), "trim");
    assert.equal(xlsx.toString(makeCell(7, "int", "A6")), "7");
    assert.equal(xlsx.toString(makeCell(null, "int?", "A7")), "");
};

const runUtilTests = () => {
    assert.equal(xlsx.isNumericKey("12"), true);
    assert.equal(xlsx.isNumericKey("-7"), true);
    assert.equal(xlsx.isNumericKey("1.25"), true);
    assert.equal(xlsx.isNumericKey("001"), false);
    assert.equal(xlsx.isNumericKey("foo"), false);

    assert.equal(xlsx.escape('a\n"b"'), 'a\\n\\"b\\"');
    assert.equal(xlsx.outdent(`\n        alpha\n        beta\n    `), "alpha\nbeta");

    assert.equal(
        xlsx.format(
            `
                hello %{name}
                  %{body}
            `,
            { name: "world", body: "line1\nline2" }
        ),
        "hello world\n  line1\n  line2"
    );
    assert.throws(() => xlsx.format("%{missing}", {}), /variable 'missing' not found/);

    const ordered = {
        "!ignore": { hidden: true },
        10: makeCell(10, "int", "A2"),
        2: makeCell(2, "int", "A3"),
        name: makeCell("Sword", "string", "A4"),
        hidden: makeCell("skip", "string", "A5"),
        none: makeCell(null, "string?", "A6"),
    } satisfies xlsx.TObject;
    assert.deepEqual(xlsx.keys(ordered, xlsx.isNotNull, ordered["!ignore"]), ["2", "10", "name"]);
    assert.deepEqual(xlsx.values(ordered, xlsx.isNotNull, ordered["!ignore"]), [
        ordered[2],
        ordered[10],
        ordered.name,
    ]);

    const enumObject = {
        "!enum": "Order",
        B: makeCell(2, "int", "B2"),
        A: makeCell(1, "int", "B3"),
    } satisfies xlsx.TObject;
    assert.deepEqual(xlsx.keys(enumObject), ["A", "B"]);

    assert.equal(xlsx.toPascalCase("task_define.row"), "TaskDefineRow");
};

const runWorkbookAndContextTests = () => {
    clearAllContexts();
    try {
        assert.throws(() => xlsx.getRunningContext(), /No running context/);

        const ctx = new xlsx.Context("context-test", "primary");
        const workbook = new xlsx.Workbook(ctx, "folder/demo.xlsx");
        ctx.add(workbook);
        assert.equal(ctx.get("demo.xlsx"), workbook);
        assert.throws(() => ctx.get("missing.xlsx"), /File not found: missing\.xlsx/);
        assert.throws(
            () => ctx.add(new xlsx.Workbook(new xlsx.Context("other", "ctx"), "other.xlsx")),
            /Context mismatch/
        );

        const duplicateNameCtx = new xlsx.Context("context-test", "secondary");
        const first = new xlsx.Workbook(duplicateNameCtx, "a/demo.xlsx");
        const second = new xlsx.Workbook(duplicateNameCtx, "b/demo.xlsx");
        duplicateNameCtx.add(first);
        duplicateNameCtx.add(second);
        assert.throws(() => duplicateNameCtx.get("demo.xlsx"), /Multiple files found:/);

        const global = xlsx.addContext(new xlsx.Context("infra", "global"));
        assert.equal(xlsx.getContext("infra", "global"), global);
        assert.throws(
            () => xlsx.addContext(new xlsx.Context("infra", "global")),
            /Context already exists/
        );
        xlsx.setRunningContext(global);
        assert.equal(xlsx.getRunningContext(), global);
        xlsx.clearRunningContext();
        assert.throws(() => xlsx.getRunningContext(), /No running context/);
        xlsx.removeContext(global);
        assert.equal(xlsx.getContext("infra", "global"), undefined);

        const sourceCtx = new xlsx.Context(xlsx.DEFAULT_WRITER, "clone-source");
        const targetCtx = new xlsx.Context("client", "clone-target");
        const sourceWorkbook = new xlsx.Workbook(sourceCtx, "clone.xlsx");
        const sourceSheet = makeSheet("main", [
            { name: "id", typename: "int", writers: ["client"] },
            { name: "client_only", typename: "string", writers: ["client"] },
            { name: "server_only", typename: "string", writers: ["server"] },
        ]);
        sourceSheet.data["!comment"] = "sheet tag";
        sourceSheet.data["1"] = makeRow({
            id: makeCell(1, "int", "A2"),
            client_only: makeCell("alpha", "string", "B2"),
            server_only: makeCell("beta", "string", "C2"),
        });
        (sourceSheet.data["1"] as xlsx.TRow)["!comment"] = "row tag";
        sourceWorkbook.add(sourceSheet);

        const cloned = sourceWorkbook.clone(targetCtx);
        const clonedSheet = cloned.get("main");
        const clonedRow = clonedSheet.data["1"] as xlsx.TRow;

        assert.deepEqual(
            clonedSheet.fields.map((field) => field.name),
            ["id", "client_only"]
        );
        assert.equal(clonedSheet.data["!comment"], "sheet tag");
        assert.equal(clonedRow["!comment"], "row tag");
        assert.equal(clonedRow.client_only.v, "alpha");
        assert.equal((clonedRow as xlsx.TObject).server_only, undefined);
        assert.notEqual(clonedRow.client_only, (sourceSheet.data["1"] as xlsx.TRow).client_only);

        assert.equal(sourceWorkbook.has("main"), true);
        sourceWorkbook.remove("main");
        assert.equal(sourceWorkbook.has("main"), false);
        assert.throws(() => sourceWorkbook.get("main"), /Sheet not found: main/);
    } finally {
        clearAllContexts();
        xlsx.clearRunningContext();
    }
};

const runRegistryTests = () => {
    const suffix = `infra_${Date.now()}`;
    withSuppressionSnapshot(() => {
        withPatchedConsole("warn", (messages) => {
            xlsx.registerType(`${suffix}_type`, (value) => value);
            xlsx.registerType(`${suffix}_type`, () => 1);
            xlsx.registerChecker(`${suffix}_checker`, () => () => true);
            xlsx.registerChecker(`${suffix}_checker`, () => () => false);
            xlsx.registerProcessor(`${suffix}_processor`, async () => {});
            xlsx.registerProcessor(`${suffix}_processor`, async () => {}, {
                stage: "after-read",
                priority: 7,
                required: true,
            });
            xlsx.registerWriter(`${suffix}_writer`, () => {});
            xlsx.registerWriter(`${suffix}_writer`, () => {});

            assert.ok(messages.some((message) => message.includes(`converter '${suffix}_type'`)));
            assert.ok(
                messages.some((message) => message.includes(`checker parser '${suffix}_checker'`))
            );
            assert.ok(
                messages.some((message) => message.includes(`processor '${suffix}_processor'`))
            );
            assert.ok(messages.some((message) => message.includes(`writer '${suffix}_writer'`)));
        });

        assert.equal(typeof xlsx.converters[`${suffix}_type`], "function");
        assert.equal(typeof xlsx.checkerParsers[`${suffix}_checker`], "function");
        assert.deepEqual(xlsx.processors[`${suffix}_processor`]?.option, {
            required: true,
            stage: "after-read",
            priority: 7,
        });
        assert.equal(typeof xlsx.writers[`${suffix}_writer`], "function");

        xlsx.suppressChecker(`${suffix}_checker`);
        xlsx.suppressProcessor(`${suffix}_processor`);
        xlsx.suppressWriter(`${suffix}_writer`);
        xlsx.suppressAllCheckers();

        assert.equal(xlsx.settings.suppressCheckers.has(`${suffix}_checker`), true);
        assert.equal(xlsx.settings.suppressProcessors.has(`${suffix}_processor`), true);
        assert.equal(xlsx.settings.suppressWriters.has(`${suffix}_writer`), true);
        assert.equal(xlsx.settings.suppressCheckers.has(xlsx.BuiltinChecker.Unique), true);
    });
};

const runParserErrorTests = () => {
    clearAllContexts();
    try {
        {
            const ctx = xlsx.addContext(new xlsx.Context(xlsx.DEFAULT_WRITER, xlsx.DEFAULT_TAG));
            const workbook = new xlsx.Workbook(ctx, "invalid-writer.xlsx");
            ctx.add(workbook);
            const rawWorkbook = {
                getSheets() {
                    return [
                        makeRawSheet("main", 1, 5, {
                            A1: "id",
                            A2: "int",
                            A3: "ghost",
                            A4: "x",
                            A5: "comment",
                        }),
                    ];
                },
            };
            assert.throws(
                () => loadHeader(workbook.path, rawWorkbook as never),
                /Writer not found: 'ghost' at A3/
            );
            clearAllContexts();
        }

        {
            const ctx = xlsx.addContext(new xlsx.Context(xlsx.DEFAULT_WRITER, xlsx.DEFAULT_TAG));
            const workbook = new xlsx.Workbook(ctx, "duplicate-field.xlsx");
            ctx.add(workbook);
            const rawWorkbook = {
                getSheets() {
                    return [
                        makeRawSheet("main", 2, 5, {
                            A1: "id",
                            B1: "id",
                            A2: "int",
                            B2: "string",
                            A3: "client",
                            B3: "client",
                            A4: "x",
                            B4: "x",
                        }),
                    ];
                },
            };
            assert.throws(
                () => loadHeader(workbook.path, rawWorkbook as never),
                /Duplicate field name: 'id' at B1/
            );
            clearAllContexts();
        }

        {
            const ctx = xlsx.addContext(new xlsx.Context(xlsx.DEFAULT_WRITER, xlsx.DEFAULT_TAG));
            const workbook = new xlsx.Workbook(ctx, "invalid-field.xlsx");
            ctx.add(workbook);
            const rawWorkbook = {
                getSheets() {
                    return [
                        makeRawSheet("main", 1, 5, {
                            A1: "bad name",
                            A2: "int",
                            A3: "client",
                            A4: "x",
                        }),
                    ];
                },
            };
            assert.throws(
                () => loadHeader(workbook.path, rawWorkbook as never),
                /Invalid field name: 'bad name' at A1/
            );
            clearAllContexts();
        }

        {
            const ctx = xlsx.addContext(new xlsx.Context(xlsx.DEFAULT_WRITER, xlsx.DEFAULT_TAG));
            const workbook = new xlsx.Workbook(ctx, "missing-checker.xlsx");
            ctx.add(workbook);
            const rawWorkbook = {
                getSheets() {
                    return [
                        makeRawSheet("main", 1, 5, {
                            A1: "id",
                            A2: "int",
                            A3: "client",
                            A4: "",
                        }),
                    ];
                },
            };
            assert.throws(
                () => loadHeader(workbook.path, rawWorkbook as never),
                /No checker defined at A4/
            );
            clearAllContexts();
        }

        {
            const ctx = xlsx.addContext(new xlsx.Context(xlsx.DEFAULT_WRITER, xlsx.DEFAULT_TAG));
            const workbook = new xlsx.Workbook(ctx, "auto-cell.xlsx");
            const sheet = makeSheet("main", [{ name: "id", typename: "auto" }]);
            workbook.add(sheet);
            ctx.add(workbook);

            const rawWorkbook = {
                getSheets() {
                    return [
                        makeRawSheet("main", 1, 6, {
                            A1: "id",
                            A6: "oops",
                        }),
                    ];
                },
            };

            assert.throws(
                () => loadBody(workbook.path, rawWorkbook as never),
                /Expected '-' at A6, but got 'oops'/
            );
            clearAllContexts();
        }

        {
            const ctx = xlsx.addContext(new xlsx.Context(xlsx.DEFAULT_WRITER, xlsx.DEFAULT_TAG));
            const workbook = new xlsx.Workbook(ctx, "missing-refer-field.xlsx");
            ctx.add(workbook);

            const rawWorkbook = {
                getSheets() {
                    return [
                        makeRawSheet("main", 1, 6, {
                            A1: "id",
                            A2: "int",
                            A3: "client",
                            A4: "@refer(rule)",
                            A6: "1",
                        }),
                    ];
                },
            };

            loadHeader(workbook.path, rawWorkbook as never);
            assert.throws(
                () => loadBody(workbook.path, rawWorkbook as never),
                /Refer field not found: rule at A4/
            );
            clearAllContexts();
        }

        {
            const ctx = xlsx.addContext(new xlsx.Context(xlsx.DEFAULT_WRITER, xlsx.DEFAULT_TAG));
            const workbook = new xlsx.Workbook(ctx, "missing-type-field.xlsx");
            ctx.add(workbook);

            const rawWorkbook = {
                getSheets() {
                    return [
                        makeRawSheet("main", 2, 6, {
                            A1: "id",
                            B1: "value",
                            A2: "int",
                            B2: "@kind",
                            A3: "client",
                            B3: "client",
                            A4: "x",
                            B4: "x",
                            A6: "1",
                            B6: "foo",
                        }),
                    ];
                },
            };

            loadHeader(workbook.path, rawWorkbook as never);
            assert.throws(
                () => loadBody(workbook.path, rawWorkbook as never),
                /Type field not found: kind at B1/
            );
            clearAllContexts();
        }
    } finally {
        clearAllContexts();
    }
};

const runPipelineFailureTests = async () => {
    const assertBuildStageFailure = async (stage: xlsx.ProcessorStage, headerOnly: boolean) => {
        clearAllContexts();
        xlsx.clearRunningContext();
        const suffix = `infra_${stage.replaceAll("-", "_")}_${Date.now()}`;
        let failure: Error | undefined;
        try {
            xlsx.registerProcessor(
                suffix,
                async () => {
                    throw new Error(`expected ${stage} failure`);
                },
                {
                    required: true,
                    stage,
                    priority: -100,
                }
            );

            try {
                await xlsx.build(["test/res/item.xlsx"], headerOnly);
            } catch (error) {
                failure = error as Error;
            }

            assert(failure);
            assert.ok(failure.message.includes(`expected ${stage} failure`));
            assert.ok(failure.message.includes(`Performing processor '${suffix}'`));
            assert.throws(() => xlsx.getRunningContext(), /No running context/);
        } finally {
            xlsx.suppressProcessor(suffix);
            clearAllContexts();
            xlsx.clearRunningContext();
        }
    };

    await assertBuildStageFailure("after-read", true);
    await assertBuildStageFailure("pre-parse", false);
    await assertBuildStageFailure("after-parse", false);
    await assertBuildStageFailure("pre-check", false);
};

const runCopyWorkbookFailureTests = () => {
    clearAllContexts();
    xlsx.clearRunningContext();
    try {
        {
            const ctx = xlsx.addContext(
                new xlsx.Context(xlsx.DEFAULT_WRITER, `copy-duplicate-${Date.now()}`)
            );
            const workbook = new xlsx.Workbook(ctx, "copy-duplicate.xlsx");
            ctx.add(workbook);

            const sheet = makeSheet("main", [
                { name: "id", typename: "int" },
                { name: "label", typename: "string" },
            ]);
            sheet.data["1"] = makeRow({
                id: makeCell(1, "int", "A2"),
                label: makeCell("first", "string", "B2"),
            });
            sheet.data["2"] = makeRow({
                id: makeCell(1, "int", "A3"),
                label: makeCell("second", "string", "B3"),
            });
            workbook.add(sheet);

            assert.throws(() => copyWorkbook(), /Duplicate key: 1, last: A2, current: A3/);
            clearAllContexts();
        }

        {
            const ctx = xlsx.addContext(
                new xlsx.Context(xlsx.DEFAULT_WRITER, `copy-empty-${Date.now()}`)
            );
            const workbook = new xlsx.Workbook(ctx, "copy-empty.xlsx");
            ctx.add(workbook);

            const sheet = makeSheet("main", [
                { name: "id", typename: "int?" },
                { name: "label", typename: "string" },
            ]);
            sheet.data["1"] = makeRow({
                id: makeCell(null, "int?", "A2"),
                label: makeCell("empty", "string", "B2"),
            });
            workbook.add(sheet);

            assert.throws(() => copyWorkbook(), /Key is empty at A2/);
            clearAllContexts();
        }
    } finally {
        clearAllContexts();
        xlsx.clearRunningContext();
    }
};

export const runInfrastructureRegressionTests = async () => {
    runErrorsAndValueTests();
    runUtilTests();
    runWorkbookAndContextTests();
    runRegistryTests();
    runParserErrorTests();
    runCopyWorkbookFailureTests();
    await runPipelineFailureTests();
};
