import * as xlsx from "../../";
let done = false;

export const makeTypename = (name: string) => {
    if (name === "items") {
        return "Items";
    }
    return name;
};

xlsx.registerProcessor(
    "workbook-typedef",
    async (workbook) => {
        if (done) {
            return;
        }
        done = true;
        xlsx.writeFile(
            "test/output/workbook-typedef.ts",
            xlsx.genWorkbookTypedef(workbook.context, (typename) => {
                if (typename === "TCell") {
                    return {
                        type: "TCell as _TCell",
                        path: "../../",
                    };
                }
                return {
                    type: makeTypename(typename),
                    path: "./client/define/index",
                };
            })
        );
    },
    {
        required: true,
        stage: "pre-parse",
    }
);
