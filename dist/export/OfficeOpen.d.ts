import { SpreadsheetDocument } from "../SpreadsheetDocument";
import { SpreadsheetOptions } from "./OpenDocument";
declare type CSpreadsheetDocument = SpreadsheetDocument & {
    _strings: string[];
};
export declare function spreadsheet(doc: CSpreadsheetDocument, options: SpreadsheetOptions): Promise<ArrayBuffer>;
export {};
