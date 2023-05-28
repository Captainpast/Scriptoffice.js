import { SpreadsheetDocument } from "./SpreadsheetDocument"
import { TextDocument } from "./TextDocument"

export interface OfficeDocumentTypes {
    "spreadsheet": SpreadsheetDocument;
    "text": TextDocument;
    "presentations": OfficeDocument;
    "graphics": OfficeDocument;
}

export class OfficeDocument {
    type: keyof OfficeDocumentTypes;
    /**the brandig will preceded by `ScriptOffice/1.0$Web_UnixLike`. the syntax should be `YourProgramm/Version$platform`*/
    generator: string = "";
    /**the last person who created the document*/
    initialCreator: string = "";
    creationDate: Date;
    /**the last person who edit the document*/
    creator: string = "";
    /**the last date where the document was edited*/
    date: Date;

    constructor() {
        this.creationDate = new Date();
    }

    async export(format: string, options?: object): Promise<any> {
        throw "not implemented"
    }
}