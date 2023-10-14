import * as OpenDocument from "./export/OpenDocument";
import * as Basic from "./export/Basic";
import * as Color from "color";
import { OfficeDocument } from "./OfficeDocument";
interface SpreadsheetDocumentExportTypes {
    "ods": [ArrayBuffer, OpenDocument.SpreadsheetOptions];
    "csv": [string, Basic.SpreadsheetCsvOptions];
    "xlsx": [ArrayBuffer, OpenDocument.SpreadsheetOptions];
}
export declare class SpreadsheetDocument extends OfficeDocument {
    sheets: SpreadsheetDocumentSheet[];
    _activeSheetIndex: number;
    constructor();
    get activeSheet(): SpreadsheetDocumentSheet;
    getSheet(identifier: number | string): SpreadsheetDocumentSheet;
    addSheet(title: string): SpreadsheetDocumentSheet;
    export<T extends keyof SpreadsheetDocumentExportTypes>(format: T, options?: SpreadsheetDocumentExportTypes[T][1]): Promise<SpreadsheetDocumentExportTypes[T][0]>;
}
export declare class SpreadsheetDocumentSheet {
    parentDocument: SpreadsheetDocument;
    cells: SpreadsheetDocumentCell[];
    title: string;
    _freezePos: {
        col: number;
        row: number;
    };
    _databases: {
        range: {
            from: {
                col: number;
                row: number;
            };
            to: {
                col: number;
                row: number;
            };
        };
    }[];
    constructor(parent: SpreadsheetDocument, title?: string);
    getCell(pos: SpreadsheetDocumentCellPosition): SpreadsheetDocumentCell;
    setCell(pos: SpreadsheetDocumentCellPosition, value: SpreadsheetDocumentCell, overwrite?: boolean): void;
    getCells(from: SpreadsheetDocumentCellPosition, to?: SpreadsheetDocumentCellPosition): SpreadsheetDocumentCell[];
    setCells(pos: SpreadsheetDocumentCellPosition, value: SpreadsheetDocumentCell): void;
    setCells(from: SpreadsheetDocumentCellPosition, to: SpreadsheetDocumentCellPosition, value: SpreadsheetDocumentCell): void;
    freezeAt(before: SpreadsheetDocumentCellPosition): boolean;
    autoFilter(range: string): boolean;
    convertToArray(from?: SpreadsheetDocumentCellPosition, to?: SpreadsheetDocumentCellPosition): SpreadsheetDocumentCellValueType[][];
}
declare type SpreadsheetDocumentCellType = "string" | "float" | "percentage" | "currency" | "date";
declare type SpreadsheetDocumentCellValueType = String | Number | Date;
export declare class SpreadsheetDocumentCell {
    parentSheet: SpreadsheetDocumentSheet;
    col: number;
    row: number;
    type: SpreadsheetDocumentCellType;
    _value: SpreadsheetDocumentCellValueType;
    _style: SpreadsheetDocumentStyle;
    constructor(col?: number, row?: number, value?: SpreadsheetDocumentCell);
    get value(): SpreadsheetDocumentCellValueType;
    set value(value: SpreadsheetDocumentCellValueType);
    get style(): SpreadsheetDocumentStyle;
    set style(value: SpreadsheetDocumentStyle);
}
declare type SpreadsheetDocumentStyleUnderline = "none" | "solid" | "wave" | "dotted" | "dash" | "long-dash" | "dot-dash" | "dot-dot-dash";
export declare class SpreadsheetDocumentStyle {
    constructor(value?: SpreadsheetDocumentStyle);
    bold: boolean;
    italic: boolean;
    underline: SpreadsheetDocumentStyleUnderline;
    /**the text underline color hex code, like `#ffffff` or default the **font-color***/
    get underlineColor(): "font-color" | string | Color;
    set underlineColor(value: "font-color" | string | Color);
    private _underlineColor;
    /**the text color hex code, like `#ffffff`*/
    get color(): string | Color;
    set color(value: string | Color);
    private _color;
    /**the cell background color hex code, like `#ffffff`*/
    get backgroundColor(): string | Color;
    set backgroundColor(value: string | Color);
    private _backgroundColor;
    /**the  width of all cells in the column*/
    columnWidth: number;
    /**the  height of all cells in the row*/
    rowHeight: number;
}
declare type SpreadsheetDocumentCellPosition = string | {
    col: string | number;
    row: number;
};
export declare function getCellName(position: {
    col: number;
    row: number;
}): string;
export {};
