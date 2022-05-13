import iRow from './iRow';

export default interface iParsedRow {
    accountName: string;
    raw: object;
    row: iRow;
}
