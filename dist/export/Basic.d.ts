import { SpreadsheetDocument } from "../SpreadsheetDocument";
export interface SpreadsheetCsvOptions {
    separator: string;
    useQuotes: boolean | string;
    sheet: string | number;
}
export declare function spreadsheetToCSV(doc: SpreadsheetDocument, options?: SpreadsheetCsvOptions): string;
