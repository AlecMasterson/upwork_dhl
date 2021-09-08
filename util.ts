import { Cell, CellFormulaValue, Column, Row, Workbook, Worksheet } from "exceljs";
import ColumnConfig from "./configs/ColumnConfig.json";
import StyleConfig from "./configs/StyleConfig.json";
import * as FileStream from "fs";

// Function for adding formatting to each column in the Worksheet.
// The first row (header) will not contain formatting.
// The column format is defined in the "ColumnConfig" JSON file.
// The format appearance is defined in the "StyleConfig" JSON file.
export function add_formatting(sheet: Worksheet): void {
    sheet.columns.forEach((column: Partial<Column>): void => {
        column.numFmt = StyleConfig["FORMATS"][ColumnConfig[column.key]?.format];
    });

    sheet.getRow(1).numFmt = null;
}

// Function for adding a table of data to the Worksheet.
// The name of the table will equal the name of the Worksheet, without spaces.
// If no rows are provided, a single row of empty values will be used.
export function add_table(sheet: Worksheet, columns: {name: string}[], rows: (number | string)[][]): void {
    const tableRows = rows.length ? rows : [columns.map((_: {name: string}): string => "")];

    sheet.addTable({
        columns,
        name: sheet.name.replace(/\s+/g, "_"),
        ref: "A1",
        rows: tableRows,
        style: { showRowStripes: true }
    });
}

// Function for adjusting each column in the Worksheet to fit the data in that column.
// The minimum column width is defined in the configuration JSON.
// A buffer of 4 is applied to each column for readability improvements.
export function adjust_column_widths(sheet: Worksheet): void {
    sheet.columns.forEach((column: Partial<Column>, _: number): void => {
        let minLength: number = StyleConfig["MIN_COLUMN_LENGTH"];

        // Get the length of the longest cell for the given column.
        column.eachCell && column.eachCell({ includeEmpty: true }, (cell: Cell, _: number): void => {
            const cellLength: number = cell.value ? cell.value.toString().length : 0;
            minLength = cellLength > minLength ? cellLength : minLength;
        });

        // Change the given column width, plus the buffer.
        column.width = minLength + 4;
    });
};

// Function for creating the directory path (if applicable) and exporting the Workbook to an output file.
// An error is thrown if any of the following is true:
// - the expected output file exists and is open
export function export_workbook(book: Workbook, path: string, fileName: string): void {
    if (!FileStream.existsSync(path)) {
        FileStream.mkdirSync(path, { recursive: true });
    }

    (async () => await book.xlsx.writeFile(`${path}/${fileName}.xlsx`))();
};

// Function for returning a specific value from the give Row as the correct type.
// An optional "type" param can be given to force the expected typel.
// An error is thrown if any of the following is true:
// - the column does not exist in the Row
// - the column is not defined in the configuration JSON
// - the expected type is not supported in the below switch statement
export function get_value(row: Row, name: string, type?: string): number | string {
    const cell: Cell = row.getCell(name);

    switch (type || ColumnConfig[name].type) {
        case "String":
            return (cell.value || "").toString();
        case "Number":
            return (cell.value as number) || 0;
        case "Formula":
            return "";
        case "Result":
            return ((cell.value as CellFormulaValue).result as number) || 0;
        default:
            throw new Error(`Column Type Not Defined: '${name}'`);
    }
}

// Function for returning a specific Worksheet from a Workbook.
// The header/key values are not retrieved automatically and must be manually set.
export function get_worksheet(book: Workbook, name: string): Worksheet {
    const sheet: Worksheet = book.getWorksheet(name);

    sheet?.columns && sheet.columns.forEach((column: Partial<Column>): void => {
        column.key = column.values[1].toString().trim();
    });

    return sheet;
}
