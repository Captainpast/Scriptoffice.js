import * as OpenDocument from "./export/OpenDocument";
import { OfficeDocument } from "./OfficeDocument"

interface SpreadsheetDocumentExportTypes {
    "ods": ArrayBuffer;
}

export class SpreadsheetDocument extends OfficeDocument {
    sheets: SpreadsheetDocumentSheet[] = [];
    _activeSheetIndex = 0;

    constructor() {
        super()
        this.type = "spreadsheet"
        this.sheets.push(new SpreadsheetDocumentSheet(this))
    }

    get activeSheet(): SpreadsheetDocumentSheet {
        if (this.sheets.length > this._activeSheetIndex) {
            return this.sheets[this._activeSheetIndex];
        } else {
            this._activeSheetIndex = 0;
            if (this.sheets.length == 0) {
                this.sheets.push(new SpreadsheetDocumentSheet(this))
            }
            return this.activeSheet;
        }
    }

    getSheet(identifier: number | string): SpreadsheetDocumentSheet {
        if (typeof(identifier) == "number" ) {
            if (this.sheets.length > identifier) {
                this._activeSheetIndex = identifier;
                return this.sheets[identifier];
            }
        } else {
            let sheet = this.sheets.find(s => s.title == identifier)
            if (sheet) {
                this._activeSheetIndex = this.sheets.indexOf(sheet);
                return sheet;
            }
        }
    }

    addSheet(title: string): SpreadsheetDocumentSheet {
        let sheet = new SpreadsheetDocumentSheet(this, title)

        this._activeSheetIndex = this.sheets.length;
        this.sheets.push(sheet)

        return sheet;
    }

    async export<T extends keyof SpreadsheetDocumentExportTypes>(format: T): Promise<SpreadsheetDocumentExportTypes[T]> {
        if (format == "ods") {
            return await OpenDocument.spreadsheet(this as any);
        } else {
            throw "not implemented"
        }
    }
}

export class SpreadsheetDocumentSheet {
    parentDocument: SpreadsheetDocument;
    cells: SpreadsheetDocumentCell[] = []
    title: string;
    _freezePos: {col: number, row: number};
    _databases: {range: {from: {col: number, row: number}, to: {col: number, row: number}}}[] = []

    constructor(parent: SpreadsheetDocument, title: string = null) {
        this.parentDocument = parent;
        if (title) {
            this.title = title;
        } else {
            this.title = "table" + (this.parentDocument.sheets.length+1);
        }
    }
    
    getCell(pos: SpreadsheetDocumentCellPosition): SpreadsheetDocumentCell {
        let {col, row} = getCellPosition(pos);
        // finde cell
        if (col != null && row != null) {
            let cell = this.cells.find(c => c.col == col && c.row == row)
            if (cell == null) {
                cell = new SpreadsheetDocumentCell(col, row);
                cell.parentSheet = this;
                this.cells.push(cell);
            }
            return cell;
        } else{
            return undefined;
        }
    }
    
    /**@deprecated WIP */
    getCells(pos: SpreadsheetDocumentCellPosition): SpreadsheetDocumentCell {
        return null;
    }

    freezeAt(before: SpreadsheetDocumentCellPosition): boolean {
        if (before) {
            this._freezePos = getCellPosition(before);
            return true;
        } else {
            this._freezePos = null;
            return false;
        }
    }

    //autoFilter(range: string): boolean;
    autoFilter(from: SpreadsheetDocumentCellPosition, to?: SpreadsheetDocumentCellPosition): boolean {
        if (from) {
            if (!to) {
                let ranges = from.toString().split(":");
                from = ranges[0];
                to = ranges[1];
            }
            let fromCell = getCellPosition(from);
            let toCell = getCellPosition(to);

            this._databases = []; // WIP
            this._databases.push({
                range: {
                    from: fromCell,
                    to: toCell
                }
            })

            return true
        } else {
            return false;
        }
    }
}

type SpreadsheetDocumentCellType = "string" | "float" | "percentage" | "currency" | "date";
type SpreadsheetDocumentCellValueType = String | Number | Date;

export class SpreadsheetDocumentCell {
    parentSheet: SpreadsheetDocumentSheet;
    col: number;
    row: number;
    type: SpreadsheetDocumentCellType;
    _value: SpreadsheetDocumentCellValueType;
    _style: SpreadsheetDocumentStyle;

    constructor(col: number = null, row: number = null) {
        this.type = "string";
        this.col = col;
        this.row = row;
    }

    get value(): SpreadsheetDocumentCellValueType { return this._value; }
    set value(value: SpreadsheetDocumentCellValueType) {
        this._value = value;
        switch (typeof(value)) {
            case "string":
                this.type = "string"
                break;
            case "number":
                this.type = "float"
                break;
            case "object":
                if (value instanceof Date) {
                    this.type = "date"
                }
                break;
        }
    }

    get style(): SpreadsheetDocumentStyle {
        this._style ??= new SpreadsheetDocumentStyle()
        return this._style;
    }
    set style(value) {
        throw "WIP";
    }
}

type SpreadsheetDocumentStyleUnderline = "none" | "solid" | "wave" | "dotted" | "dash" | "long-dash" | "dot-dash" | "dot-dot-dash"; // WIP: "bold", "double"
export class SpreadsheetDocumentStyle {
    bold = false;
    italic = false;
    underline: SpreadsheetDocumentStyleUnderline = "none";
    /**the text underline color hex code, like `#ffffff` or default the **font-color***/
    underlineColor: "font-color" | string = "font-color";
    /**the text color hex code, like `#ffffff`*/
    color: string;
    /**the cell background color hex code, like `#ffffff`*/
    backgroundColor: string;

    /**the  width of all cells in the column*/
    columnWidth: number;
    /**the  height of all cells in the row*/
    rowHeight: number;
}


type SpreadsheetDocumentCellPosition = string | { col: string | number, row: number };
function getCellPosition(pos: SpreadsheetDocumentCellPosition) {
    let col: number, row: number;
    if (typeof(pos) == "string") {
        let res = pos.match(/([A-Za-z]+)(\d+)/)
        if (res != null) {
            col = getValueFromCol(res[1])
            row = parseInt(res[2])
        }
    } else {
        if (typeof(pos.col) == "string") {
            col = getValueFromCol(pos.col);
        } else {
            col = pos.col;
        }
        row = pos.row;
    }

    return {col, row}
}

const alphabet = "abcdefghijklmnopqrstuvwxyz".split("")
/**
 * gets a column `id` as string of characters between **A-Z** and returns the matching number value
 * @example getValueFromCol("AB") => 28
 * 
 * @param id the identifier for the column
 * @returns the number value of the column
 */
function getValueFromCol(id: string): number {
    id = id.toLowerCase();
    let value = 0;
    for (let i = 0; i < id.length; i++) {
        const char = alphabet.indexOf(id[i]);
        if (char > -1) {
            value += (char +1) * Math.pow(26, id.length -1 - i);
        }
    }
    return value;
}

/** 
 * converts an value to a number system between **A-Z**
 * @example getColFromValue(28) => "AB"
 * 
 * @param value the number value for the column
 * @returns the string identifier of the column
 */
function getColFromValue(value: number): string {
    let id = "";
    while (value > 0) {
        const modulo = (value - 1) % 26;
        id = alphabet[modulo] + id;
        value = Math.floor((value - modulo) / 26);
    }
    return id.toUpperCase()
}

export function getCellName(position: { col: number, row: number }) {
    return getColFromValue(position.col) + position.row
}