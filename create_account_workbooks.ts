import ExcelJS, { Row, Workbook, Worksheet } from "exceljs";
import * as CsvWriter from "csv-writer";
import AccountMapping from "./configs/AccountMapping";
import ColumnsExportData from "./configs/output_columns/ExportData.json";
import ColumnsExportDest from "./configs/output_columns/ExportDestinationCharges.json";
import ColumnsImportData from "./configs/output_columns/ImportData.json";
import ColumnsImportDest from "./configs/output_columns/ImportDestinationCharges.json";
import StyleConfig from "./configs/StyleConfig.json";
import SummaryCSV from "./configs/output_columns/SummaryCSV.json";
import { add_formatting, add_table, adjust_column_widths, export_workbook, get_value, get_worksheet } from "./util";
import Lodash from "lodash";

/* =================================================
====================================================
CUSTOM TYPES
====================================================
================================================= */

interface ColumnConfig {
    format?: string;
    name: string;
    type?: string;
}

interface Summary {
    "Account": string;
    "Invoice Type": string;
    "Total Charge": number;
}

interface WorksheetConfig {
    accountId: string;
    columns: ColumnConfig[];
    sheet: Worksheet;
}

/* =================================================
====================================================
FUNCTIONS
====================================================
================================================= */

// For each account in the AccountMapping configuration, do the following:
// - Create a new Workbook for the given account.
// - Attempt to create the four different Worksheets from the input data.
// - If any of the new Worksheets were created for the current account, do the following:
//   - Create the summary Worksheet based on the output data.
//   - Format the Worksheets in the new Workbook (excluding the Summary Worksheet).
//   - Export the new Workbook to an output file.
export function create_account_workbooks(book: Workbook, date: string): void {
    const summary: {[accountName: string]: {[sheetName: string]: number}} = {};

    Object.keys(AccountMapping).forEach((accountName: string): void => {
        const newBook: Workbook = new ExcelJS.Workbook();

        [
            {accountId: AccountMapping[accountName].exportId, columns: ColumnsExportData, sheet: get_worksheet(book, "Export Data")},
            {accountId: AccountMapping[accountName].exportId, columns: ColumnsExportDest, sheet: get_worksheet(book, "Export Destination Charges")},
            {accountId: AccountMapping[accountName].importId, columns: ColumnsImportData, sheet: get_worksheet(book, "Import Data")},
            {accountId: AccountMapping[accountName].importId, columns: ColumnsImportDest, sheet: get_worksheet(book, "Import Destination Charges")}
        ]
        .filter((config: WorksheetConfig): boolean => config.sheet !== undefined)
        .forEach((config: WorksheetConfig): void => {
            const totalCharge: number = create_worksheet(newBook, accountName, config);
            Lodash.set(summary, `${accountName}.${config.sheet.name}`, totalCharge);
        });

        if (newBook.worksheets.length > 0) {
            create_summary_worksheet(newBook);

            newBook.worksheets.forEach((sheet: Worksheet): void => {
                if (sheet.name !== "Summary") {
                    add_formatting(get_worksheet(newBook, sheet.name));
                }

                adjust_column_widths(sheet);
            });

            export_workbook(newBook, "results", `${accountName}_${date}_shipment_report`);
        }
    });

    const rows: Summary[] = [];
    Object.keys(summary).forEach((accountName: string): void => {
        Object.keys(summary[accountName])
            .filter((sheetName: string): boolean => summary[accountName][sheetName] !== 0)
            .forEach((sheetName: string): void => {
                rows.push({"Account": accountName, "Invoice Type": sheetName, "Total Charge": summary[accountName][sheetName]});
            });
    });

    CsvWriter.createObjectCsvWriter({
        header: SummaryCSV.map((column: {name: string}): string => column.name),
        path: `results/${date}_summary.csv`
    }).writeRecords(rows);
}

// Function to create the Worksheet that summarizes the other Worksheets.
// All formulas utilize the table and column names from the other Worksheets.
// Apply the "Currency" format to all values in the Worksheet.
function create_summary_worksheet(book: Workbook): void {
    book.addWorksheet("Summary").addRows([
        ["Export Data", {formula: "IFERROR(SUM(Export_Data[Total Charge]), 0)"}],
        ["Export Destination Charges", {formula: "IFERROR(SUM(Export_Destination_Charges[Total Charge]), 0)"}],
        ["Import Data", {formula: "IFERROR(SUM(Import_Data[Total Charge]), 0)"}],
        ["Import Destination Charges", {formula: "IFERROR(SUM(Import_Destination_Charges[Total Charge]), 0)"}],
        ["Total", {formula: "SUM(B1:B4)"}],
        ["Total if by Credit Card", {formula: "B5*1.03"}]
    ]);

    get_worksheet(book, "Summary").getColumn(2).numFmt = StyleConfig["FORMATS"]["Currency"];
}

function create_worksheet(book: Workbook, accountName: string, config: WorksheetConfig): number {
    // Create the empty list of rows needed for the table. It will be populated below.
    const rows: (number | string)[][] = [];

    // Track the sum of the "Total Charge" column for this table.
    let totalTotalCharge: number = 0;

    // For each row in this Worksheet:
    // - only utilize rows where the "Billing Account" matches the current account ID
    // - map the row to all defined columns in the configuration JSON file
    // - apply few exceptions to the values "Account Name" and "Total Charge"
    config.sheet.eachRow((row: Row, _: number): void => {
        if (get_value(row, "Billing Account") == config.accountId) {
            rows.push(config.columns.map((column: ColumnConfig): number | string => {
                switch (column.name) {
                    case "Account Name":
                        return accountName;
                    case "Total Charge":
                        let totalCharge: number = get_value(row, "Grand Total") as number;
                        if (!config.sheet.name.includes("Destination")) {
                            totalCharge *= AccountMapping[accountName].markup;
                        }

                        totalTotalCharge += totalCharge;
                        return totalCharge;
                    default:
                        return get_value(row, column.name, column.type);
                }
            }));
        }
    });

    // If data exists, create the new Worksheet and insert the table of data associated with the current account ID.
    rows.length > 0 && add_table(book.addWorksheet(config.sheet.name), config.columns, rows);

    // Return the sum of the "Total Charge" column.
    return totalTotalCharge;
}
