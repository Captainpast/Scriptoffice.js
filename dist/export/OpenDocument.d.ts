import { SpreadsheetDocument, SpreadsheetDocumentStyle } from "../SpreadsheetDocument";
declare type CSpreadsheetDocument = SpreadsheetDocument & {
    _styles: CSpreadsheetDocumentStyle[];
};
declare type CSpreadsheetDocumentStyle = SpreadsheetDocumentStyle & {
    _name: string;
    _type: string;
    _target: number;
    _stringified: string;
};
export declare function spreadsheet(doc: CSpreadsheetDocument): Promise<ArrayBuffer>;
export {};
