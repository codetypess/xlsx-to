import * as xlsx from "../../index.js";

xlsx.registerStringifyRule("task", (workbook: xlsx.Workbook) => {
    return xlsx.simpleSheets(workbook);
});
