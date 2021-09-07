import { Row, Workbook, Worksheet } from "exceljs";
import AccountMapping from "./configs/AccountMapping";
import SummaryColumns from "./configs/SummaryColumns.json";
import { get_value, get_worksheet } from "./util";
import * as CsvWriter from "csv-writer";
import Lodash from "lodash";

// Create a CSV output file summarizing the input Workbook content.
// The columns for the CSV can be found in the SummaryColumns.json configuration file.
export function create_summary_csv(book: Workbook, date: string): void {
    CsvWriter.createObjectCsvWriter({
        header: SummaryColumns.map((column: {name: string}): string => column.name),
        path: `results/${date}_summary.csv`
    }).writeRecords(get_rows(book));
}

function get_rows(book: Workbook): any[] {
    let rows = [];

    book.worksheets.map((sheet: Worksheet): string => sheet.name).filter((sheetName: string): boolean => sheetName !== "Summary").forEach((sheetName: string): void => {
        const sheetSummary = {};

        const get_account_name = (billingAccount: string): string => {
            const field: string = sheetName.includes("Export") ? "exportId" : "importId";
            return Object.keys(AccountMapping).find((accountName: string): boolean => AccountMapping[accountName][field] == billingAccount);
        };

        get_worksheet(book, sheetName).eachRow((row: Row, _: number): void => {
            const billingAccount: string = get_value(row, "Billing Account", "String") as string;
            const accountName: string = get_account_name(billingAccount);

            if (accountName) {
                Lodash.set(sheetSummary, `${accountName}.Account`, accountName);
                Lodash.set(sheetSummary, `${accountName}.Invoice Type`, sheetName);

                const grandTotal: number = get_value(row, "Grand Total", "Result") as number;
                const totalCharge: number = get_value(row, "Total Charge", "Number") as number;
                Lodash.update(sheetSummary, `${accountName}.Grand Total`, (val: number): number => val ? val + grandTotal : grandTotal);
                Lodash.update(sheetSummary, `${accountName}.Total Charge`, (val: number): number => val ? val + totalCharge : totalCharge);
            }
        });

        rows = rows.concat(Object.values(sheetSummary));
    });

    return rows;
}
