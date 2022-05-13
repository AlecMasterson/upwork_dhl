import * as Path from 'path';
import * as CsvWriter from 'csv-writer';
import * as FS from 'fs';
import {flatten, groupBy, sum} from 'lodash';
import AccountMap from './Accounts';
import iParsedRow from './types/iParsedRow';
import iResponse from './types/iResponse';
import iRow from './types/iRow';
import {exportAccount, readFilePath} from './util';

// From the command arguments, obtain the directory path containing all the CSV files.
if (process.argv.length !== 3) throw new Error('REQUIRED: 1 Argument, the Directory Path');
const dirPath: string = process.argv[2];

// Verify that the provided directory path is a true directory.
if (!FS.lstatSync(dirPath).isDirectory()) {
    throw new Error(`Invalid Directory: ${dirPath}\n`);
}

// Create the output directory if it doesn't exist.
const today: Date = new Date();
const outputDir: string = `./results/${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
if (!FS.existsSync(outputDir)) {
    FS.mkdirSync(outputDir, {recursive: true});
}

// Obtain the file paths to all CSV files in the provided directory.
// Asynchronously read each CSV file.
const streams: Array<Promise<iResponse>> = FS.readdirSync(dirPath)
    .filter((fileName: string): boolean => Path.extname(fileName) === '.csv')
    .map((fileName: string): string => Path.join(dirPath, fileName))
    .map(readFilePath);

(async (): Promise<void> => {
    // Wait for each asynchronous reading to have completed.
    // Flatten the list of iParsedRow objects into a final list of objects to parse.
    const responses: iResponse[] = flatten(await Promise.all(streams));
    const data: iParsedRow[] = flatten(responses.map((response: iResponse): iParsedRow[] => response.parsedRows));
    const headers: string[] = responses[0].headers;

    // The list to contain the overall summarized rows;
    const summary: iRow[] = [];

    // Group by the "accountName" and process each grouping separately.
    const groupedAccountName: {[accountName: string]: iParsedRow[]} = groupBy(data, 'accountName');
    Object.keys(groupedAccountName).forEach((accountName: string): void => {
        // Get all the rows for the given account.
        const rows: iRow[] = flatten(groupedAccountName[accountName].map((parsedRow: iParsedRow): iRow => parsedRow.row));

        // Get all the raw rows for the given account.
        const raw: object[] = flatten(groupedAccountName[accountName].map((parsedRow: iParsedRow): object => parsedRow.raw));

        // Group all rows for the given account by the "Invoice Number" and the "Product Name".
        // Each grouping will be used for a separate aggregation.
        const groupedInvoiceNumber: {[invoiceNumber: string]: iRow[]} = groupBy(rows, 'Invoice Number');
        const groupedProductName: {[productName: string]: iRow[]} = groupBy(rows, 'Product Name');

        // Use the "Invoice Number" grouping to perform an aggregation of the "Total Charge".
        // This aggregation will create the overall summarized row for that account and invoice.
        Object.keys(groupedInvoiceNumber).forEach((invoiceNumber: string): void => {
            // Calculate the aggregated "Total Charge" of the given "Invoice Number".
            const totalCharge: number = groupedInvoiceNumber[invoiceNumber]
                .reduce((total: number, row: iRow): number => total + parseFloat(row['Total Charge'] as unknown as string), 0);

            summary.push({
                'Billing Account': AccountMap[accountName].id as string,
                'Due Date': groupedInvoiceNumber[invoiceNumber][0]['Due Date'],
                'Invoice Date': groupedInvoiceNumber[invoiceNumber][0]['Invoice Date'],
                'Invoice Number': invoiceNumber,
                'Product Name': groupedInvoiceNumber[invoiceNumber][0]['Product Name'],
                'Total Charge': totalCharge
            });
        });

        // Use the "Product Name" grouping to perform an aggregation of the "Total Charge".
        // This aggregation will be used for the individual account's CSV file for the given account.
        const productNameCharges: Array<any[]> = Object.keys(groupedProductName)
            .map((productName: string): any[] => {
                // Calculate the aggregated "Total Charge" of the given "Product Name".
                const totalCharge: number = groupedProductName[productName]
                    .reduce((total: number, row: iRow): number => total + parseFloat(row['Total Charge'] as unknown as string), 0);

                return [productName, totalCharge];
            });

        const total: number = sum(productNameCharges.map((summaryRow: any[]): number => summaryRow[1] as number));
        productNameCharges.push(['Total', total]);

        exportAccount(Path.join(outputDir, accountName), headers, productNameCharges, raw);
    });

    CsvWriter.createObjectCsvWriter({
        header: [
            {id: 'Billing Account', title: 'Billing Account'},
            {id: 'Due Date', title: 'Due Date'},
            {id: 'Invoice Date', title: 'Invoice Date'},
            {id: 'Invoice Number', title: 'Invoice Number'},
            {id: 'Product Name', title: 'Product Name'},
            {id: 'Total Charge', title: 'Total Charge'}
        ],
        path: './results/summary.csv'
    }).writeRecords(summary).then((): void => console.log('Summary Has Been Saved!'));
})();
