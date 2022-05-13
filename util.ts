import CsvParser from 'csv-parser';
import ExcelJS, {Cell, Column, Workbook, Worksheet} from 'exceljs';
import * as FS from 'fs';
import {get} from 'lodash';
import iParsedRow from './types/iParsedRow';
import iResponse from './types/iResponse';
import AccountMap from './Accounts';

// Function to adjust the column widths of an Excel sheet.
// It will adjust the column to the longest text for that column (with a buffer).
function adjustColumnWidths(sheet: Worksheet): void {
    sheet.columns.forEach((column: Partial<Column>): void => {
        let minLength: number = 10;

        column.eachCell && column.eachCell({includeEmpty: true}, (cell: Cell): void => {
            const cellLength: number = cell.value ? cell.value.toString().length : 0;
            minLength = Math.max(minLength, cellLength);
        });

        column.width = minLength + 4;
    });
}

// Function to export an individual account to an Excel file.
// The Excel file will contain 2 sheets, the "Summary" and "Data".
// The "Summary" will have the "per Product Name" breakdown.
// The "Data" will have the original raw shipment data.
export function exportAccount(outputPath: string, headers: string[], summary: Array<any[]>, raw: object[]): void {
    const book: Workbook = new ExcelJS.Workbook();

    book.addWorksheet('Summary').addRows([['Product Name', 'Total Charges'], ...summary]);
    adjustColumnWidths(book.getWorksheet('Summary'));

    const rawArray: Array<any[]> = raw.map((row: object): any[] => headers.map((header: string): any => get(row, header, '')));
    book.addWorksheet('Data').addRows([headers, ...rawArray]);
    adjustColumnWidths(book.getWorksheet('Data'));

    book.xlsx.writeFile(`${outputPath}.xlsx`);
}

// Function to change the "Billing Account" of a given row to the new mapping.
// It will only alter the row if the AccountMap has a key that matches one of the below columns.
// If there is no new mapping defined, an error will be thrown.
function mapAccount(filePath: string, row: any): iParsedRow {
    const newAccount: string | undefined = Object.keys(AccountMap)
        .find((accountName: string): boolean => (
            row['Senders Name']?.toUpperCase() === accountName.toUpperCase() ||
            row['Sender Contact']?.toUpperCase() === accountName.toUpperCase() ||
            row['Receivers Name']?.toUpperCase() === accountName.toUpperCase()
        ));

    if (newAccount === undefined) {
        throw new Error(
            `Failed to Find Mapping for '${filePath}'\n` +
            `Senders Name:\t${row['Senders Name']}\n` +
            `Sender Contact:\t${row['Sender Contact']}\n` +
            `Receivers Name:\t${row['Receivers Name']}\n`
        );
    }

    return {
        accountName: newAccount,
        raw: row,
        row: {...row, 'Billing Account': AccountMap[newAccount].id as string}
    };
}

// Function to asynchronously read a file path using the CSV parser library.
export async function readFilePath(filePath: string): Promise<iResponse> {
    return new Promise((resolve: any, reject: any): void => {
        let headers: string[] = [];
        const parsedRows: iParsedRow[] = [];

        FS.createReadStream(filePath)
            .pipe(CsvParser())
            .on('data', (row: any): void => {
                parsedRows.push(mapAccount(filePath, row));
            })
            .on('headers', (csvHeaders: string[]): void => {
                headers = csvHeaders;
            })
            .on('end', (): void => resolve({headers, parsedRows}))
            .on('error', (error: unknown): void => reject(`Failed to Read File: ${filePath}\n${error}\n`));
    });
}
