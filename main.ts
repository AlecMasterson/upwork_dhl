import ExcelJS, { Workbook } from "exceljs";
import { create_account_workbooks } from "./create_account_workbooks";
import { create_summary_csv } from "./create_summary_csv";

/* =================================================
====================================================
CONFIGURATION
====================================================
================================================= */

const INPUT_FILE_DATE: string = "6-28-21";
const INPUT_FILE_NAME: string = `data/examples/PF - ${INPUT_FILE_DATE}.xlsx`;

/* =================================================
====================================================
MAIN PROCESS
====================================================
================================================= */

// Read the input file as a new Workbook, then do the following:
// - create all XLSX Workbooks for each applicable account
// - create the summary CSV file
new ExcelJS.Workbook().xlsx.readFile(INPUT_FILE_NAME).then((book: Workbook): void => {
    create_account_workbooks(book, INPUT_FILE_DATE);
    create_summary_csv(book, INPUT_FILE_DATE);
});
