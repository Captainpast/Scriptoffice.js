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
export interface SpreadsheetOptions {
    compressionLevel: number;
}
export declare function escapeXML(value: any): string;
export declare function dateToString(date: Date): string;
export declare function spreadsheet(doc: CSpreadsheetDocument, options: SpreadsheetOptions): Promise<ArrayBuffer>;
export {};
