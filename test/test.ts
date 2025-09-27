import * as xlsx from "..";

import "./init";

const t = Date.now();

const OUTPUT_DIR = "test/output";

const defines: string[] = ["export * from './custom';"];

const makeTypename = (name: string) => {
    if (name === "items") {
        return "Items";
    }
    return name;
};

xlsx.registerWriter("client", (path, data, processor) => {
    if (processor === "define") {
        const name = xlsx.toPascalCase(`${xlsx.filename(path)}_${data["!name"]}`);
        const marshal = `export const ${name} = `;
        xlsx.writeFile(
            `${OUTPUT_DIR}/client/define/${name}.ts`,
            xlsx.stringifyTs(data, { indent: 4, marshal })
        );
        defines.push(`export * from "./${name}";`);
    } else if (processor === "stringify") {
        const name = xlsx.filename(path);
        xlsx.writeFile(
            `${OUTPUT_DIR}/client/data/${name}.json`,
            xlsx.stringifyJson(data, { indent: 2 })
        );
    } else if (processor === "typedef") {
        const name = xlsx.filename(path);
        const types = xlsx.genTsTypedef(path, "client", (typename) => {
            return {
                type: makeTypename(typename),
                path: "../define/index",
            };
        });
        xlsx.writeFile(`${OUTPUT_DIR}/client/types/${name}.ts`, types);
    } else {
        throw new Error(`Unknown handler processor: ${processor}`);
    }
});

xlsx.registerWriter("server", (path, data, processor) => {
    if (processor === "define") {
        const name = `${xlsx.filename(path)}_${data["!name"]}`;
        const marshal = `return `;
        xlsx.writeFile(
            `${OUTPUT_DIR}/server/define/${name}.lua`,
            xlsx.stringifyLua(data, { indent: 4, marshal })
        );
    } else if (processor === "stringify") {
        const name = xlsx.filename(path);
        const marshal = `return `;
        xlsx.writeFile(
            `${OUTPUT_DIR}/server/data/${name}.lua`,
            xlsx.stringifyLua(data, { indent: 2, marshal })
        );
    } else if (processor === "typedef") {
        const name = xlsx.filename(path);
        const types = xlsx.genLuaTypedef(path, "server", (typename) => {
            return { type: makeTypename(typename) };
        });
        xlsx.writeFile(
            `${OUTPUT_DIR}/server/types/${name}.lua`,
            xlsx.outdent(`
                -- AUTO GENERATED, DO NOT MODIFY!
                
                ${types}
            `)
        );
    } else {
        throw new Error(`Unknown handler processor: ${processor}`);
    }
});

await xlsx.parse(["test/res/item.xlsx", "test/res/task.xlsx"]);

xlsx.writeFile(
    "test/output/workbook-typedef.ts",
    xlsx.genWorkbookTypedef((typename) => {
        return {
            type: makeTypename(typename),
            path: "./client/define/index",
        };
    })
);
xlsx.writeFile("test/output/client/define/index.ts", defines.join("\n"));

console.log(Date.now() - t);
