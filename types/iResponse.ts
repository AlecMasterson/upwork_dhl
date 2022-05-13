import iParsedRow from './iParsedRow';

export default interface iResponse {
    headers: string[];
    parsedRows: iParsedRow[];
}
