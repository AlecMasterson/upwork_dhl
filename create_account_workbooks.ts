import ExcelJS, { Column, Row, Workbook, Worksheet } from "exceljs";
import AccountMapping from "./configs/AccountMapping";
import StyleConfig from "./configs/StyleConfig.json";
import { add_formatting, add_table, adjust_column_widths, export_workbook, get_value, get_worksheet } from "./util";

// For each account in the AccountMapping configuration, do the following:
// - Create a new Workbook for the given account.
// - Create the four different Worksheets from the input data.
// - If any of the Worksheets had data for the given account, do the following:
//   - Create the summary Worksheet based on the output data (excluding the Summary Worksheet).
//   - Format the Worksheets in the new Workbook.
//   - Export the new Workbook to an output file.
export function create_account_workbooks(book: Workbook, date: string): void {
    Object.keys(AccountMapping).forEach((accountName: string): void => {
        const newBook: Workbook = new ExcelJS.Workbook();

        const hasData: boolean[] = [
            create_worksheet(book, newBook, "Export Data", accountName, AccountMapping[accountName].exportId),
            create_worksheet(book, newBook, "Export Destination Charges", accountName, AccountMapping[accountName].exportId),
            create_worksheet(book, newBook, "Import Data", accountName, AccountMapping[accountName].importId),
            create_worksheet(book, newBook, "Import Destination Charges", accountName, AccountMapping[accountName].importId)
        ];

        if (hasData.includes(true)) {
            add_formulas(newBook, accountName);
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
}

/* =================================================
====================================================
FUNCTIONS
====================================================
================================================= */

// Function for adding the necessary formulas to the Worksheets.
function add_formulas(book: Workbook, accountName: string): void {
    const get_range = (column: Column): string => {
        return `${column.letter}2:${column.letter}${column.values.length - 1}`;
    };

    const data_formula = (name: string): void => {
        const sheet: Worksheet = get_worksheet(book, name);
        if (sheet) {
            const cellLoc: string = `${sheet.getColumn("Grand Total").letter}2`;
            const formula: string = `IF(${cellLoc}="","",${cellLoc}*${AccountMapping[accountName].markup})`;
            sheet.fillFormula(get_range(sheet.getColumn("Total Charge")), formula);
        }
    };

    const destination_formula = (name: string): void => {
        const sheet: Worksheet = get_worksheet(book, name);
        if (sheet) {
            const formula: string = `${sheet.getColumn("Grand Total").letter}2`;
            sheet.fillFormula(get_range(sheet.getColumn("Total Charge")), formula);
        }
    };

    data_formula("Export Data");
    destination_formula("Export Destination Charges");
    data_formula("Import Data");
    destination_formula("Import Destination Charges");
}

// Function to create the Worksheet that summarizes the other Worksheets.
// All formulas utilize the table and column names from the other Worksheets.
// Apply the "Currency" format to all values in the Worksheet.
function create_summary_worksheet(book: Workbook): void {
    book.addWorksheet("Summary").addRows([
        ["Export Data", {formula: "SUM(Export_Data[Total Charge])"}],
        ["Export Destination Charges", {formula: "SUM(Export_Destination_Charges[Total Charge])"}],
        ["Import Data", {formula: "SUM(Import_Data[Total Charge])"}],
        ["Import Destination Charges", {formula: "SUM(Import_Destination_Charges[Total Charge])"}],
        ["Total", {formula: "SUM(B1:B4)"}],
        ["Total if by Credit Card", {formula: "B5*1.03"}]
    ]);

    get_worksheet(book, "Summary").getColumn(2).numFmt = StyleConfig["FORMATS"]["Currency"];
}

function create_worksheet(book: Workbook, newBook: Workbook, sheetName: string, accountName: string, accountId: string): boolean {
    // Get the Worksheet from the input file that the output file will be based on.
    // If the input file does not exist, return false.
    const sheet: Worksheet = get_worksheet(book, sheetName);
    if (!sheet) {
        return false;
    }

    // Get the column names/headers found in this Worksheet, removing any duplicates.
    // These can be matched with the keys found in the ColumnConfig.json configuration file.
    let columns: string[] = sheet.columns.map((column: Partial<Column>): string => column.key);
    columns = columns.filter((column: string, i: number): boolean => columns.indexOf(column) === i);

    // Create the empty list of rows needed for the table. It will be populated later.
    // Create the list of columns needed for the table. It must follow the below type.
    let rows: (number | string)[][] = [];
    let tableColumns: {name: string}[] = columns.map((column: string): {name: string} => ({name: column}));

    // For each row in this Worksheet:
    // - only utilize rows where the "Billing Account" matches the current account ID
    // - map the row to all previosuly defined columns in this Worksheet
    sheet.eachRow((row: Row, _: number): void => {
        if (get_value(row, "Billing Account") == accountId) {
            rows.push(columns.map((column: string): number | string => get_value(row, column)));
        }
    });

    // Add the "Account Name" column to the front of this Worksheet.
    rows = rows.map((row: (number | string)[]) => ([accountName] as (number | string)[]).concat(row));
    tableColumns = [{name: "Account Name"}].concat(tableColumns);

    // Add a new Worksheet and insert the table of data associated with the current account ID.
    add_table(newBook.addWorksheet(sheetName), tableColumns, rows);

    // Return whether or not there was data in the table.
    // This will be used later to determine if the account should be exported to an output file.
    return rows.length > 0;
}
